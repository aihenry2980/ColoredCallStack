using EnvDTE;
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
        private static readonly Guid TextEditorFontCategory = new Guid("A27B4E24-A735-4D1D-B8E7-9716E1E3D8E0");
        private static readonly DrawingColor DefaultLightNamespaceColor = DrawingColor.FromArgb(255, 90, 96, 104);
        private static readonly DrawingColor DefaultLightFunctionColor = DrawingColor.FromArgb(255, 0, 82, 153);
        private static readonly DrawingColor DefaultLightParamNameColor = DrawingColor.FromArgb(255, 122, 63, 0);
        private static readonly DrawingColor DefaultLightParamValueColor = DrawingColor.FromArgb(255, 0, 100, 0);
        private static readonly DrawingColor DefaultLightFileColor = DrawingColor.FromArgb(255, 28, 98, 139);
        private static readonly DrawingColor DefaultLightLineColor = DrawingColor.FromArgb(255, 170, 75, 0);
        private static readonly DrawingColor DefaultLightPunctuationColor = DrawingColor.FromArgb(255, 33, 37, 41);
        private static readonly DrawingColor DefaultDarkNamespaceColor = DrawingColor.FromArgb(255, 156, 163, 175);
        private static readonly DrawingColor DefaultDarkFunctionColor = DrawingColor.FromArgb(255, 86, 156, 214);
        private static readonly DrawingColor DefaultDarkParamNameColor = DrawingColor.FromArgb(255, 206, 145, 120);
        private static readonly DrawingColor DefaultDarkParamValueColor = DrawingColor.FromArgb(255, 181, 206, 168);
        private static readonly DrawingColor DefaultDarkFileColor = DrawingColor.FromArgb(255, 156, 220, 254);
        private static readonly DrawingColor DefaultDarkLineColor = DrawingColor.FromArgb(255, 255, 198, 109);
        private static readonly DrawingColor DefaultDarkPunctuationColor = DrawingColor.FromArgb(255, 208, 208, 208);

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

        private void DebuggerEvents_OnEnterBreakMode(dbgEventReason Reason, ref dbgExecutionAction ExecutionAction)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ShowToolWindow();
            RefreshCallStack();
        }

        private void DebuggerEvents_OnContextChanged(EnvDTE.Process NewProcess, EnvDTE.Program NewProgram, EnvDTE.Thread NewThread, EnvDTE.StackFrame NewStackFrame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            RefreshCallStack();
        }

        private void DebuggerEvents_OnEnterRunMode(dbgEventReason Reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ClearCallStack();
        }

        private void DebuggerEvents_OnEnterDesignMode(dbgEventReason Reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
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
                    if (_dte.ActiveDocument?.Selection is TextSelection selection)
                    {
                        selection.GotoLine(lineNumber, true);
                    }
                }
            }
            catch
            {
                // Best-effort navigation; ignore failures.
            }

            RefreshCallStack();
        }

        private void RefreshCallStack()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var control = GetToolWindowControl(create: true);
            if (control == null || _dte?.Debugger == null)
            {
                return;
            }

            ApplyPalette(control);
            ApplyDisplayOptions(control);
            ApplyThemeMode(control);
            ApplyTextEditorFont(control);

            EnvDTE.Thread thread = _dte.Debugger.CurrentThread;
            if (thread == null)
            {
                control.ClearCallStack();
                return;
            }

            var frames = new List<StackFrame>();
            foreach (StackFrame frame in thread.StackFrames)
            {
                frames.Add(frame);
            }

            control.UpdateCallStack(frames, _dte.Debugger.CurrentStackFrame);
        }

        private void ClearCallStack()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var control = GetToolWindowControl(create: false);
            control?.ClearCallStack();
        }

        private void ShowToolWindow()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ToolWindowPane window = FindToolWindow(typeof(ColoredCallStack), 0, true);
            if (window?.Frame is IVsWindowFrame windowFrame)
            {
                Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(windowFrame.Show());
            }
        }

        private ColoredCallStackControl GetToolWindowControl(bool create)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ToolWindowPane window = FindToolWindow(typeof(ColoredCallStack), 0, create);
            if (window?.Content is ColoredCallStackControl control)
            {
                if (_callStackControl == null)
                {
                    _callStackControl = control;
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

                return _callStackControl;
            }

            return null;
        }

        private void Options_ThemeModeChanged(object sender, EventArgs e)
        {
            _ = JoinableTaskFactory.RunAsync(async () =>
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
                RefreshCallStack();
            });
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
            bool showParameterTypes = options?.ShowParameterTypes ?? false;
            bool showLineNumbers = true;
            bool showFilePath = false;
            bool hexDisplayMode = _dte?.Debugger != null && _dte.Debugger.HexDisplayMode;

            control.SetDisplayOptions(showParameterTypes, showLineNumbers, showFilePath, hexDisplayMode);
        }

        private void Control_DisplayOptionsChanged(object sender, ColoredCallStackControl.DisplayOptionsChangedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var options = GetDialogPage(typeof(ColoredCallStackOptions)) as ColoredCallStackOptions;
            if (options == null)
            {
                return;
            }

            bool refreshNeeded = false;
            if (e.ShowParameterTypes.HasValue)
            {
                options.ShowParameterTypes = e.ShowParameterTypes.Value;
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
                RefreshCallStack();
            }
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
                ToMediaColor(NormalizeColor(options.LightNamespaceColor, DefaultLightNamespaceColor)),
                ToMediaColor(NormalizeColor(options.LightFunctionColor, DefaultLightFunctionColor)),
                ToMediaColor(NormalizeColor(options.LightParamNameColor, DefaultLightParamNameColor)),
                ToMediaColor(NormalizeColor(options.LightParamValueColor, DefaultLightParamValueColor)),
                ToMediaColor(NormalizeColor(options.LightFileColor, DefaultLightFileColor)),
                ToMediaColor(NormalizeColor(options.LightLineColor, DefaultLightLineColor)),
                ToMediaColor(NormalizeColor(options.LightPunctuationColor, DefaultLightPunctuationColor)));

            var dark = new ColoredCallStackControl.Palette(
                ToMediaColor(NormalizeColor(options.DarkNamespaceColor, DefaultDarkNamespaceColor)),
                ToMediaColor(NormalizeColor(options.DarkFunctionColor, DefaultDarkFunctionColor)),
                ToMediaColor(NormalizeColor(options.DarkParamNameColor, DefaultDarkParamNameColor)),
                ToMediaColor(NormalizeColor(options.DarkParamValueColor, DefaultDarkParamValueColor)),
                ToMediaColor(NormalizeColor(options.DarkFileColor, DefaultDarkFileColor)),
                ToMediaColor(NormalizeColor(options.DarkLineColor, DefaultDarkLineColor)),
                ToMediaColor(NormalizeColor(options.DarkPunctuationColor, DefaultDarkPunctuationColor)));

            control.SetPalettes(light, dark);
        }

        private void ApplyTextEditorFont(ColoredCallStackControl control)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (control == null)
            {
                return;
            }

            if (TryGetTextEditorFont(out string family, out double size))
            {
                control.SetFont(family, size);
            }
        }

        private static MediaColor ToMediaColor(DrawingColor color)
        {
            return MediaColor.FromArgb(color.A, color.R, color.G, color.B);
        }

        private static DrawingColor NormalizeColor(DrawingColor value, DrawingColor fallback)
        {
            return value.IsEmpty ? fallback : value;
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

            if (TryGetFileInfoFromDebugFrame(frame, out fileName, out lineNumber))
            {
                return true;
            }

            try
            {
                dynamic dyn = frame;
                fileName = dyn.FileName as string;
                lineNumber = (int)dyn.LineNumber;
                return !string.IsNullOrEmpty(fileName) && lineNumber > 0;
            }
            catch
            {
                fileName = null;
                lineNumber = 0;
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

            return !string.IsNullOrEmpty(fileName) && lineNumber > 0;
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
