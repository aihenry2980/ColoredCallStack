using EnvDTE;
using EnvDTE90a;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Debugger.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Text.RegularExpressions;
using System.Reflection;
using Task = System.Threading.Tasks.Task;
using DrawingColor = System.Drawing.Color;
using MediaColor = System.Windows.Media.Color;

namespace ColorCallStack
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(ColorCallStackPackage.PackageGuidString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideToolWindow(typeof(ColoredCallStack))]
    [ProvideOptionPage(typeof(ColoredCallStackOptions), "ColorCallStack", "General", 0, 0, true)]
    [ProvideProfile(typeof(ColoredCallStackOptions), "ColorCallStack", "General", 0, 0, true)]
    [ProvideAutoLoad(UIContextGuids80.Debugging, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class ColorCallStackPackage : AsyncPackage
    {
        /// <summary>
        /// ColorCallStackPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "a0ed6e99-f242-48ff-a531-1eb668125a8e";

        private DTE2 _dte;
        private DebuggerEvents _debuggerEvents;
        private ColoredCallStackControl _callStackControl;
        private bool _frameActivatedHooked;
        private bool _displayOptionsHooked;
        private bool _fontSizeHooked;
        private bool _resetHooked;
        private bool _fontSizeStepsLoaded;
        private int _fontSizeSteps;
        private int _emptyBreakRefreshRetries;
        private CancellationTokenSource _refreshCts;
        private const int ContextRefreshDebounceMs = 120;
        private const int ContextRefreshMinGapMs = 250;
        private const int BreakFollowupRefreshMs = 90;
        private const int BreakGuaranteedRefreshMs = 260;
        private const int RetryRefreshDebounceMs = 60;
        private const int MaxEmptyBreakRefreshRetries = 40;
        private const double FontStepDip = 1.0;
        private const int MinFontSizeSteps = -8;
        private const int MaxFontSizeSteps = 30;
        private static readonly Guid TextEditorFontCategory = new Guid("A27B4E24-A735-4D1D-B8E7-9716E1E3D8E0");
        private DateTime _lastRenderedAtUtc = DateTime.MinValue;

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            _dte = await GetServiceAsync(typeof(SDTE)) as DTE2;
            Assumes.Present(_dte);
            _debuggerEvents = _dte.Events.DebuggerEvents;
            _debuggerEvents.OnEnterBreakMode += DebuggerEvents_OnEnterBreakMode;
            _debuggerEvents.OnContextChanged += DebuggerEvents_OnContextChanged;
            _debuggerEvents.OnEnterRunMode += DebuggerEvents_OnEnterRunMode;
            _debuggerEvents.OnEnterDesignMode += DebuggerEvents_OnEnterDesignMode;

            ColoredCallStackOptions.ThemeModeChanged += Options_ThemeModeChanged;

            await ColoredCallStackCommand.InitializeAsync(this);
        }

        internal void RefreshCallStackForUserRequest()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            CancelPendingRefresh();
            RefreshCallStack(onlyIfVisible: false);
        }

        private void DebuggerEvents_OnEnterBreakMode(dbgEventReason Reason, ref dbgExecutionAction ExecutionAction)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _emptyBreakRefreshRetries = 0;
            CancelPendingRefresh();
            bool rendered = RefreshCallStackAndReport(onlyIfVisible: false);
            if (!rendered)
            {
                ScheduleBreakWarmupRefreshes();
            }
        }

        private void DebuggerEvents_OnContextChanged(EnvDTE.Process NewProcess, EnvDTE.Program NewProgram, EnvDTE.Thread NewThread, EnvDTE.StackFrame NewStackFrame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_dte?.Debugger?.CurrentMode != dbgDebugMode.dbgBreakMode)
            {
                return;
            }

            if ((DateTime.UtcNow - _lastRenderedAtUtc).TotalMilliseconds < ContextRefreshMinGapMs)
            {
                return;
            }

            QueueRefreshCallStack(onlyIfVisible: true, delayMs: ContextRefreshDebounceMs);
        }

        private void DebuggerEvents_OnEnterRunMode(dbgEventReason Reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            CancelPendingRefresh();
            _emptyBreakRefreshRetries = 0;
            _lastRenderedAtUtc = DateTime.MinValue;
            // Keep last rendered frames while stepping to avoid UI churn on every F10.
        }

        private void DebuggerEvents_OnEnterDesignMode(dbgEventReason Reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            CancelPendingRefresh();
            _emptyBreakRefreshRetries = 0;
            _lastRenderedAtUtc = DateTime.MinValue;
            ClearCallStack();
        }

        private void Control_FrameActivated(object sender, StackFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_dte?.Debugger == null || frame == null)
            {
                return;
            }

            if (_dte.Debugger.CurrentMode != dbgDebugMode.dbgBreakMode)
            {
                return;
            }

            try
            {
                _dte.Debugger.CurrentStackFrame = frame;
            }
            catch
            {
                return;
            }

            try
            {
                if (TryGetFileInfo(frame, out string fileName, out int lineNumber))
                {
                    _dte.ItemOperations.OpenFile(fileName);
                    if (lineNumber > 0 && _dte.ActiveDocument?.Selection is TextSelection selection)
                    {
                        selection.GotoLine(lineNumber, true);
                    }
                }
            }
            catch
            {
                // Best-effort navigation; ignore failures.
            }

            RefreshCallStack(onlyIfVisible: false);
        }

        private void QueueRefreshCallStack(bool onlyIfVisible = true, int delayMs = ContextRefreshDebounceMs)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            CancelPendingRefresh();
            _refreshCts = new CancellationTokenSource();
            CancellationToken token = _refreshCts.Token;

            JoinableTaskFactory.RunAsync(async () =>
            {
                await Task.Delay(delayMs, token);
                await JoinableTaskFactory.SwitchToMainThreadAsync(token);
                RefreshCallStack(onlyIfVisible);
            }).FileAndForget("ColorCallStack/DebouncedRefresh");
        }

        private void CancelPendingRefresh()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_refreshCts != null)
            {
                _refreshCts.Cancel();
                _refreshCts.Dispose();
                _refreshCts = null;
            }
        }

        private void ScheduleBreakWarmupRefreshes()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            JoinableTaskFactory.RunAsync(async () =>
            {
                await Task.Delay(BreakFollowupRefreshMs);
                await JoinableTaskFactory.SwitchToMainThreadAsync();
                if (_dte?.Debugger?.CurrentMode == dbgDebugMode.dbgBreakMode)
                {
                    if (RefreshCallStackAndReport(onlyIfVisible: false))
                    {
                        return;
                    }
                }

                await Task.Delay(BreakGuaranteedRefreshMs);
                await JoinableTaskFactory.SwitchToMainThreadAsync();
                if (_dte?.Debugger?.CurrentMode == dbgDebugMode.dbgBreakMode)
                {
                    RefreshCallStack(onlyIfVisible: false);
                }
            }).FileAndForget("ColorCallStack/BreakWarmupRefreshes");
        }

        private void RefreshCallStack(bool onlyIfVisible)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            RefreshCallStackAndReport(onlyIfVisible);
        }

        private bool RefreshCallStackAndReport(bool onlyIfVisible)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var control = GetToolWindowControl(create: false);
            if (control == null && _dte?.Debugger?.CurrentMode == dbgDebugMode.dbgBreakMode && IsToolWindowVisible())
            {
                control = GetToolWindowControl(create: true);
            }

            if (control == null || _dte?.Debugger == null)
            {
                if (_dte?.Debugger?.CurrentMode == dbgDebugMode.dbgBreakMode &&
                    IsToolWindowVisible() &&
                    _emptyBreakRefreshRetries < MaxEmptyBreakRefreshRetries)
                {
                    _emptyBreakRefreshRetries++;
                    QueueRefreshCallStack(onlyIfVisible: false, delayMs: RetryRefreshDebounceMs);
                }

                return false;
            }

            if (onlyIfVisible && !IsToolWindowVisible())
            {
                return false;
            }

            ApplyDisplayOptions(control);

            StackFrame currentFrame = null;
            try
            {
                currentFrame = _dte.Debugger.CurrentStackFrame;
            }
            catch
            {
                currentFrame = null;
            }

            EnvDTE.Thread thread = _dte.Debugger.CurrentThread;
            if (thread == null)
            {
                if (currentFrame != null)
                {
                    _emptyBreakRefreshRetries = 0;
                    control.UpdateCallStack(new List<StackFrame> { currentFrame }, currentFrame);
                    _lastRenderedAtUtc = DateTime.UtcNow;
                    return true;
                }

                if (_dte.Debugger.CurrentMode == dbgDebugMode.dbgBreakMode &&
                    _emptyBreakRefreshRetries < MaxEmptyBreakRefreshRetries)
                {
                    _emptyBreakRefreshRetries++;
                    QueueRefreshCallStack(onlyIfVisible: false, delayMs: RetryRefreshDebounceMs);
                    return false;
                }

                _emptyBreakRefreshRetries = 0;
                control.ClearCallStack();
                return false;
            }

            var frames = new List<StackFrame>();
            try
            {
                foreach (StackFrame frame in thread.StackFrames)
                {
                    frames.Add(frame);
                }
            }
            catch
            {
                if (_dte.Debugger.CurrentMode == dbgDebugMode.dbgBreakMode &&
                    _emptyBreakRefreshRetries < MaxEmptyBreakRefreshRetries)
                {
                    _emptyBreakRefreshRetries++;
                    QueueRefreshCallStack(onlyIfVisible: false, delayMs: RetryRefreshDebounceMs);
                    return false;
                }

                _emptyBreakRefreshRetries = 0;
                control.ClearCallStack();
                return false;
            }

            if (frames.Count == 0 && currentFrame != null)
            {
                frames.Add(currentFrame);
            }

            if (frames.Count == 0 && _dte.Debugger.CurrentMode == dbgDebugMode.dbgBreakMode)
            {
                if (_emptyBreakRefreshRetries < MaxEmptyBreakRefreshRetries)
                {
                    _emptyBreakRefreshRetries++;
                    QueueRefreshCallStack(onlyIfVisible: false, delayMs: RetryRefreshDebounceMs);
                    return false;
                }
            }

            _emptyBreakRefreshRetries = 0;
            control.UpdateCallStack(frames, currentFrame);
            _lastRenderedAtUtc = DateTime.UtcNow;
            return frames.Count > 0;
        }

        private bool IsToolWindowVisible()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ToolWindowPane window = FindToolWindow(typeof(ColoredCallStack), 0, false);
            if (!(window?.Frame is IVsWindowFrame frame))
            {
                return false;
            }

            return frame.IsVisible() == VSConstants.S_OK;
        }

        private void ClearCallStack()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var control = GetToolWindowControl(create: false);
            control?.ClearCallStack();
        }

        private ColoredCallStackControl GetToolWindowControl(bool create)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ToolWindowPane window = FindToolWindow(typeof(ColoredCallStack), 0, create);
            if (window?.Content is ColoredCallStackControl control)
            {
                bool newControl = !ReferenceEquals(_callStackControl, control);
                if (newControl)
                {
                    _callStackControl = control;
                    _frameActivatedHooked = false;
                    _displayOptionsHooked = false;
                    _fontSizeHooked = false;
                    _resetHooked = false;
                }

                if (!_frameActivatedHooked)
                {
                    _callStackControl.FrameActivated += Control_FrameActivated;
                    _frameActivatedHooked = true;
                }
                if (!_displayOptionsHooked)
                {
                    _callStackControl.DisplayOptionsChanged += Control_DisplayOptionsChanged;
                    _displayOptionsHooked = true;
                }
                if (!_fontSizeHooked)
                {
                    _callStackControl.FontSizeStepRequested += Control_FontSizeStepRequested;
                    _fontSizeHooked = true;
                }
                if (!_resetHooked)
                {
                    _callStackControl.ResetRequested += Control_ResetRequested;
                    _resetHooked = true;
                }

                if (newControl)
                {
                    ApplyPalette(_callStackControl);
                    ApplyDisplayOptions(_callStackControl);
                    ApplyThemeMode(_callStackControl);
                    ApplyTextEditorFont(_callStackControl);
                }

                return _callStackControl;
            }

            return null;
        }

        private void Options_ThemeModeChanged(object sender, EventArgs e)
        {
            JoinableTaskFactory.RunAsync(async () =>
            {
                await JoinableTaskFactory.SwitchToMainThreadAsync();
                var control = GetToolWindowControl(create: false);
                if (control == null)
                {
                    return;
                }

                ApplyPalette(control);
                ApplyDisplayOptions(control);
                ApplyThemeMode(control);
                ApplyTextEditorFont(control);
                RefreshCallStack(onlyIfVisible: false);
            }).FileAndForget("ColorCallStack/ApplyOptions");
        }

        private void ApplyThemeMode(ColoredCallStackControl control)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (control == null)
            {
                return;
            }

            var options = GetDialogPage(typeof(ColoredCallStackOptions)) as ColoredCallStackOptions;
            CallStackThemeMode mode = options?.ThemeMode ?? CallStackThemeMode.Auto;
            control.SetThemeMode(mode);
        }

        private void ApplyDisplayOptions(ColoredCallStackControl control)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (control == null)
            {
                return;
            }

            var options = GetDialogPage(typeof(ColoredCallStackOptions)) as ColoredCallStackOptions;
            bool showNamespace = options?.ShowNamespace ?? true;
            bool showParameterTypes = options?.ShowParameterTypes ?? false;
            bool showLineNumbers = options?.ShowLineNumbers ?? true;
            bool showFilePath = options?.ShowFilePath ?? true;
            bool hexDisplayMode = _dte?.Debugger != null && _dte.Debugger.HexDisplayMode;

            control.SetDisplayOptions(showNamespace, showParameterTypes, showLineNumbers, showFilePath, hexDisplayMode);
        }

        private void Control_DisplayOptionsChanged(object sender, ColoredCallStackControl.DisplayOptionsChangedEventArgs e)
        {
            if (!ThreadHelper.CheckAccess())
            {
                JoinableTaskFactory.RunAsync(async () =>
                {
                    await JoinableTaskFactory.SwitchToMainThreadAsync();
                    Control_DisplayOptionsChangedCore(e);
                }).FileAndForget("ColorCallStack/DisplayOptionsChanged");
                return;
            }

            Control_DisplayOptionsChangedCore(e);
        }

        private void Control_DisplayOptionsChangedCore(ColoredCallStackControl.DisplayOptionsChangedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var options = GetDialogPage(typeof(ColoredCallStackOptions)) as ColoredCallStackOptions;
            if (options == null)
            {
                return;
            }

            bool refreshNeeded = false;
            if (e.ShowNamespace.HasValue)
            {
                options.ShowNamespace = e.ShowNamespace.Value;
                refreshNeeded = true;
            }

            if (e.ShowParameterTypes.HasValue)
            {
                options.ShowParameterTypes = e.ShowParameterTypes.Value;
                refreshNeeded = true;
            }

            if (e.ShowLineNumbers.HasValue)
            {
                options.ShowLineNumbers = e.ShowLineNumbers.Value;
                refreshNeeded = true;
            }

            if (e.ShowFilePath.HasValue)
            {
                options.ShowFilePath = e.ShowFilePath.Value;
                refreshNeeded = true;
            }

            if (e.HexDisplayMode.HasValue && _dte?.Debugger != null)
            {
                _dte.Debugger.HexDisplayMode = e.HexDisplayMode.Value;
                refreshNeeded = true;
            }

            options.SaveSettingsToStorage();

            var control = GetToolWindowControl(create: false);
            if (control != null)
            {
                ApplyDisplayOptions(control);
            }

            if (refreshNeeded)
            {
                CancelPendingRefresh();
                RefreshCallStack(onlyIfVisible: false);
            }
        }

        private void Control_FontSizeStepRequested(object sender, int stepDelta)
        {
            if (!ThreadHelper.CheckAccess())
            {
                JoinableTaskFactory.RunAsync(async () =>
                {
                    await JoinableTaskFactory.SwitchToMainThreadAsync();
                    Control_FontSizeStepRequestedCore(stepDelta);
                }).FileAndForget("ColorCallStack/FontSizeStepRequested");
                return;
            }

            Control_FontSizeStepRequestedCore(stepDelta);
        }

        private void Control_FontSizeStepRequestedCore(int stepDelta)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (stepDelta == 0)
            {
                return;
            }

            EnsureFontSizeStepsLoaded();
            _fontSizeSteps = ClampInt(_fontSizeSteps + stepDelta, MinFontSizeSteps, MaxFontSizeSteps);
            PersistFontSizeSteps();
            var control = GetToolWindowControl(create: false);
            if (control != null)
            {
                ApplyTextEditorFont(control);
            }
        }

        private void Control_ResetRequested(object sender, EventArgs e)
        {
            if (!ThreadHelper.CheckAccess())
            {
                JoinableTaskFactory.RunAsync(async () =>
                {
                    await JoinableTaskFactory.SwitchToMainThreadAsync();
                    Control_ResetRequestedCore();
                }).FileAndForget("ColorCallStack/ResetRequested");
                return;
            }

            Control_ResetRequestedCore();
        }

        private void Control_ResetRequestedCore()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var options = GetDialogPage(typeof(ColoredCallStackOptions)) as ColoredCallStackOptions;
            if (options == null)
            {
                return;
            }

            options.ResetToFactoryDefaults();
            options.SaveSettingsToStorage();
            _fontSizeStepsLoaded = false;

            var control = GetToolWindowControl(create: false);
            if (control != null)
            {
                ApplyPalette(control);
                ApplyDisplayOptions(control);
                ApplyThemeMode(control);
                ApplyTextEditorFont(control);
            }

            CancelPendingRefresh();
            RefreshCallStack(onlyIfVisible: false);
        }

        private void ApplyPalette(ColoredCallStackControl control)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (control == null)
            {
                return;
            }

            var options = GetDialogPage(typeof(ColoredCallStackOptions)) as ColoredCallStackOptions;
            if (options == null)
            {
                return;
            }

            var light = new ColoredCallStackControl.Palette(
                ToMediaColor(NormalizeColor(options.LightNamespaceColor, ColoredCallStackDefaultColors.LightNamespace)),
                ToMediaColor(NormalizeColor(options.LightFunctionColor, ColoredCallStackDefaultColors.LightFunction)),
                ToMediaColor(NormalizeColor(options.LightParamNameColor, ColoredCallStackDefaultColors.LightParamName)),
                ToMediaColor(NormalizeColor(options.LightParamValueColor, ColoredCallStackDefaultColors.LightParamValue)),
                ToMediaColor(NormalizeColor(options.LightFileColor, ColoredCallStackDefaultColors.LightFile)),
                ToMediaColor(NormalizeColor(options.LightLineColor, ColoredCallStackDefaultColors.LightLine)),
                ToMediaColor(NormalizeColor(options.LightPunctuationColor, ColoredCallStackDefaultColors.LightPunctuation)));

            var dark = new ColoredCallStackControl.Palette(
                ToMediaColor(NormalizeColor(options.DarkNamespaceColor, ColoredCallStackDefaultColors.DarkNamespace)),
                ToMediaColor(NormalizeColor(options.DarkFunctionColor, ColoredCallStackDefaultColors.DarkFunction)),
                ToMediaColor(NormalizeColor(options.DarkParamNameColor, ColoredCallStackDefaultColors.DarkParamName)),
                ToMediaColor(NormalizeColor(options.DarkParamValueColor, ColoredCallStackDefaultColors.DarkParamValue)),
                ToMediaColor(NormalizeColor(options.DarkFileColor, ColoredCallStackDefaultColors.DarkFile)),
                ToMediaColor(NormalizeColor(options.DarkLineColor, ColoredCallStackDefaultColors.DarkLine)),
                ToMediaColor(NormalizeColor(options.DarkPunctuationColor, ColoredCallStackDefaultColors.DarkPunctuation)));

            control.SetPalettes(light, dark);
        }

        private void ApplyTextEditorFont(ColoredCallStackControl control)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (control == null)
            {
                return;
            }

            EnsureFontSizeStepsLoaded();
            string family = "Consolas";
            double size = 12.0;
            TryGetTextEditorFont(out family, out size);

            double adjustedSize = Math.Max(1.0, size + (_fontSizeSteps * FontStepDip));
            control.SetFont(family, adjustedSize);

            var options = GetDialogPage(typeof(ColoredCallStackOptions)) as ColoredCallStackOptions;
            double namespaceScale = PercentToScale(options?.NamespaceFontSizePercent ?? 100);
            double functionScale = PercentToScale(options?.FunctionFontSizePercent ?? 100);
            double parameterScale = PercentToScale(options?.ParameterFontSizePercent ?? 100);
            double lineScale = PercentToScale(options?.LineFontSizePercent ?? 100);
            double fileScale = PercentToScale(options?.FileFontSizePercent ?? 100);

            control.SetTokenFontScales(namespaceScale, functionScale, parameterScale, lineScale, fileScale);
            control.SetTokenFontFamilies(
                options?.NamespaceFontFace,
                options?.FunctionFontFace,
                options?.ParameterFontFace,
                options?.LineFontFace,
                options?.FileFontFace);
        }

        private void EnsureFontSizeStepsLoaded()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_fontSizeStepsLoaded)
            {
                return;
            }

            var options = GetDialogPage(typeof(ColoredCallStackOptions)) as ColoredCallStackOptions;
            _fontSizeSteps = ClampInt(options?.FontSizeAdjustmentSteps ?? 0, MinFontSizeSteps, MaxFontSizeSteps);
            _fontSizeStepsLoaded = true;
        }

        private void PersistFontSizeSteps()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var options = GetDialogPage(typeof(ColoredCallStackOptions)) as ColoredCallStackOptions;
            if (options == null || options.FontSizeAdjustmentSteps == _fontSizeSteps)
            {
                return;
            }

            options.FontSizeAdjustmentSteps = _fontSizeSteps;
            options.SaveSettingsToStorage();
        }

        private static MediaColor ToMediaColor(DrawingColor color)
        {
            return MediaColor.FromArgb(color.A, color.R, color.G, color.B);
        }

        private static DrawingColor NormalizeColor(DrawingColor value, DrawingColor fallback)
        {
            return value.IsEmpty ? fallback : value;
        }

        private static int ClampInt(int value, int min, int max)
        {
            if (value < min)
            {
                return min;
            }

            if (value > max)
            {
                return max;
            }

            return value;
        }

        private static double PercentToScale(int percent)
        {
            int clamped = ClampInt(percent, 50, 300);
            return clamped / 100.0;
        }

        private bool TryGetTextEditorFont(out string fontFamily, out double fontSize)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            fontFamily = "Consolas";
            fontSize = 12.0;

            var storage = GetService(typeof(SVsFontAndColorStorage)) as IVsFontAndColorStorage;
            if (storage == null)
            {
                return false;
            }

            Guid category = TextEditorFontCategory;
            int hr = storage.OpenCategory(ref category, (uint)__FCSTORAGEFLAGS.FCSF_READONLY);
            if (ErrorHandler.Failed(hr))
            {
                return false;
            }

            try
            {
                var logFont = new LOGFONTW[1];
                var info = new FontInfo[1];
                hr = storage.GetFont(logFont, info);
                if (ErrorHandler.Failed(hr))
                {
                    return false;
                }

                if (!string.IsNullOrWhiteSpace(info[0].bstrFaceName))
                {
                    fontFamily = info[0].bstrFaceName;
                }

                if (info[0].wPointSize > 0)
                {
                    fontSize = info[0].wPointSize * 96.0 / 72.0;
                }

                return true;
            }
            finally
            {
                storage.CloseCategory();
            }
        }

        private static bool TryGetFileInfo(StackFrame frame, out string fileName, out int lineNumber)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            fileName = null;
            lineNumber = 0;
            if (frame == null)
            {
                return false;
            }

            if (TryGetFileInfoFromStackFrame2(frame, out fileName, out lineNumber))
            {
                return true;
            }

            if (TryGetStringPropertyValue(frame, "FileName", out string candidateFile))
            {
                fileName = candidateFile;
            }

            if (TryGetIntPropertyValue(frame, "LineNumber", out int candidateLine))
            {
                lineNumber = candidateLine;
            }

            if (!string.IsNullOrEmpty(fileName))
            {
                return true;
            }

            if (TryGetFileInfoFromDebugFrame(frame, out fileName, out lineNumber))
            {
                return true;
            }
            return false;
        }

        private static bool TryGetFileInfoFromStackFrame2(StackFrame frame, out string fileName, out int lineNumber)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            fileName = null;
            lineNumber = 0;
            if (!(frame is StackFrame2 frame2))
            {
                return false;
            }

            try
            {
                fileName = frame2.FileName;
            }
            catch
            {
                fileName = null;
            }

            try
            {
                lineNumber = unchecked((int)frame2.LineNumber);
            }
            catch
            {
                lineNumber = 0;
            }

            return !string.IsNullOrEmpty(fileName);
        }

        private static bool TryGetStringPropertyValue(object source, string propertyName, out string value)
        {
            value = null;
            if (source == null || string.IsNullOrEmpty(propertyName))
            {
                return false;
            }

            try
            {
                PropertyInfo property = source.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
                if (property == null || !property.CanRead)
                {
                    return false;
                }

                object raw = property.GetValue(source);
                value = raw as string ?? raw?.ToString();
                return !string.IsNullOrEmpty(value);
            }
            catch
            {
                // Fall back to IDispatch-based access for COM-backed RCWs.
            }

            try
            {
                object raw = source.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase,
                    binder: null,
                    target: source,
                    args: null);

                value = raw as string ?? raw?.ToString();
                return !string.IsNullOrEmpty(value);
            }
            catch
            {
                value = null;
                return false;
            }
        }

        private static bool TryGetIntPropertyValue(object source, string propertyName, out int value)
        {
            value = 0;
            if (source == null || string.IsNullOrEmpty(propertyName))
            {
                return false;
            }

            try
            {
                PropertyInfo property = source.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
                if (property == null || !property.CanRead)
                {
                    return false;
                }

                object raw = property.GetValue(source);
                switch (raw)
                {
                    case int intValue:
                        value = intValue;
                        return true;
                    case short shortValue:
                        value = shortValue;
                        return true;
                    case long longValue when longValue <= int.MaxValue && longValue >= int.MinValue:
                        value = (int)longValue;
                        return true;
                    case uint uintValue when uintValue <= int.MaxValue:
                        value = (int)uintValue;
                        return true;
                    default:
                        return int.TryParse(raw?.ToString(), out value);
                }
            }
            catch
            {
                // Fall back to IDispatch-based access for COM-backed RCWs.
            }

            try
            {
                object raw = source.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase,
                    binder: null,
                    target: source,
                    args: null);

                switch (raw)
                {
                    case int intValue:
                        value = intValue;
                        return true;
                    case short shortValue:
                        value = shortValue;
                        return true;
                    case long longValue when longValue <= int.MaxValue && longValue >= int.MinValue:
                        value = (int)longValue;
                        return true;
                    case uint uintValue when uintValue <= int.MaxValue:
                        value = (int)uintValue;
                        return true;
                    default:
                        return int.TryParse(raw?.ToString(), out value);
                }
            }
            catch
            {
                value = 0;
                return false;
            }
        }

        private static bool TryGetFileInfoFromDebugFrame(StackFrame frame, out string fileName, out int lineNumber)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            fileName = null;
            lineNumber = 0;

            if (!TryGetDebugFrame(frame, out IDebugStackFrame2 debugFrame))
            {
                return false;
            }

            try
            {
                if (TryGetFileInfoFromFrameInfo(debugFrame, out fileName, out lineNumber))
                {
                    return true;
                }

                if (!TryGetDocumentContext(debugFrame, out IDebugDocumentContext2 context))
                {
                    return false;
                }

                return TryGetFileInfoFromDocumentContext(context, out fileName, out lineNumber);
            }
            catch
            {
                fileName = null;
                lineNumber = 0;
                return false;
            }
        }

        private static bool TryGetFileInfoFromFrameInfo(IDebugStackFrame2 debugFrame, out string fileName, out int lineNumber)
        {
            fileName = null;
            lineNumber = 0;
            if (debugFrame == null)
            {
                return false;
            }

            try
            {
                var info = new FRAMEINFO[1];
                var flags = enum_FRAMEINFO_FLAGS.FIF_FUNCNAME | enum_FRAMEINFO_FLAGS.FIF_FUNCNAME_LINES;
                int hr = debugFrame.GetInfo(flags, 10, info);
                if (!Microsoft.VisualStudio.ErrorHandler.Succeeded(hr) || info == null || info.Length == 0)
                {
                    return false;
                }

                string funcName = info[0].m_bstrFuncName;
                return TryParseFileAndLineFromText(funcName, out fileName, out lineNumber);
            }
            catch
            {
                fileName = null;
                lineNumber = 0;
                return false;
            }
        }

        private static bool TryGetDocumentContext(IDebugStackFrame2 debugFrame, out IDebugDocumentContext2 context)
        {
            context = null;
            if (debugFrame == null)
            {
                return false;
            }

            int hr = debugFrame.GetDocumentContext(out context);
            if (Microsoft.VisualStudio.ErrorHandler.Succeeded(hr) && context != null)
            {
                return true;
            }

            hr = debugFrame.GetCodeContext(out IDebugCodeContext2 codeContext);
            if (!Microsoft.VisualStudio.ErrorHandler.Succeeded(hr) || codeContext == null)
            {
                return false;
            }

            hr = codeContext.GetDocumentContext(out context);
            return Microsoft.VisualStudio.ErrorHandler.Succeeded(hr) && context != null;
        }

        private static bool TryGetFileInfoFromDocumentContext(IDebugDocumentContext2 context, out string fileName, out int lineNumber)
        {
            fileName = null;
            lineNumber = 0;
            if (context == null)
            {
                return false;
            }

            string fullName = null;
            if (Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetName(enum_GETNAME_TYPE.GN_FILENAME, out fullName)))
            {
                fileName = fullName;
            }

            if (string.IsNullOrEmpty(fileName) && Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetName(enum_GETNAME_TYPE.GN_URL, out string urlName)))
            {
                fileName = urlName;
            }
            if (string.IsNullOrEmpty(fileName) && Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetName(enum_GETNAME_TYPE.GN_BASENAME, out string baseName)))
            {
                fileName = baseName;
            }
            if (string.IsNullOrEmpty(fileName) && Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetName(enum_GETNAME_TYPE.GN_NAME, out string name)))
            {
                fileName = name;
            }
            if (string.IsNullOrEmpty(fileName) && Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetName(enum_GETNAME_TYPE.GN_MONIKERNAME, out string moniker)))
            {
                fileName = moniker;
            }
            if (string.IsNullOrEmpty(fileName) && Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetName(enum_GETNAME_TYPE.GN_TITLE, out string title)))
            {
                fileName = title;
            }

            var begin = new TEXT_POSITION[1];
            var end = new TEXT_POSITION[1];
            if (Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetStatementRange(begin, end)))
            {
                uint lineValue = begin[0].dwLine;
                if (lineValue != uint.MaxValue)
                {
                    lineNumber = unchecked((int)lineValue) + 1;
                }
            }
            if (lineNumber <= 0 && Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetSourceRange(begin, end)))
            {
                uint lineValue = begin[0].dwLine;
                if (lineValue != uint.MaxValue)
                {
                    lineNumber = unchecked((int)lineValue) + 1;
                }
            }

            return !string.IsNullOrEmpty(fileName);
        }

        private static bool TryParseFileAndLineFromText(string text, out string fileName, out int lineNumber)
        {
            fileName = null;
            lineNumber = 0;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            Match match = Regex.Match(text, @"(?<file>[A-Za-z]:\\[^:\r\n\)\]]+\.\w+)\((?<line>\d+)\)");
            if (!match.Success)
            {
                match = Regex.Match(text, @"(?<file>[A-Za-z]:\\[^:\r\n\)\]]+\.\w+):(?<line>\d+)");
            }
            if (!match.Success)
            {
                match = Regex.Match(text, @"(?<file>\\\\[^:\r\n\)\]]+\.\w+)\((?<line>\d+)\)");
            }
            if (!match.Success)
            {
                match = Regex.Match(text, @"(?<file>\\\\[^:\r\n\)\]]+\.\w+):(?<line>\d+)");
            }
            if (!match.Success)
            {
                match = Regex.Match(text, @"(?<file>[^\\/:()\r\n\]]+\.\w+)\((?<line>\d+)\)");
            }
            if (!match.Success)
            {
                match = Regex.Match(text, @"(?<file>[^\\/:()\r\n\]]+\.\w+):(?<line>\d+)");
            }

            if (match.Success)
            {
                fileName = match.Groups["file"].Value;
                int.TryParse(match.Groups["line"].Value, out lineNumber);
                return !string.IsNullOrEmpty(fileName);
            }

            match = Regex.Match(text, @"(?<file>[A-Za-z]:\\[^:\r\n\)\]]+\.\w+)");
            if (!match.Success)
            {
                match = Regex.Match(text, @"(?<file>\\\\[^:\r\n\)\]]+\.\w+)");
            }
            if (!match.Success)
            {
                match = Regex.Match(text, @"(?<file>[^\\/:()\r\n\]]+\.\w+)");
            }
            if (match.Success)
            {
                fileName = match.Groups["file"].Value;
                return !string.IsNullOrEmpty(fileName);
            }

            return false;
        }

        private static bool TryGetDebugFrame(StackFrame frame, out IDebugStackFrame2 debugFrame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            debugFrame = null;
            if (frame == null)
            {
                return false;
            }

            if (frame is IDebugStackFrame2 direct)
            {
                debugFrame = direct;
                return true;
            }

            IntPtr unk = IntPtr.Zero;
            IntPtr ppv = IntPtr.Zero;
            try
            {
                unk = Marshal.GetIUnknownForObject(frame);
                Guid iid = typeof(IDebugStackFrame2).GUID;
                int hr = Marshal.QueryInterface(unk, ref iid, out ppv);
                if (!Microsoft.VisualStudio.ErrorHandler.Succeeded(hr) || ppv == IntPtr.Zero)
                {
                    return false;
                }

                debugFrame = (IDebugStackFrame2)Marshal.GetObjectForIUnknown(ppv);
                return debugFrame != null;
            }
            catch
            {
                debugFrame = null;
                return false;
            }
            finally
            {
                if (ppv != IntPtr.Zero)
                {
                    Marshal.Release(ppv);
                }
                if (unk != IntPtr.Zero)
                {
                    Marshal.Release(unk);
                }
            }
        }

        #endregion
    }
}

