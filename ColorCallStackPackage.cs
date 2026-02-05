using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

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
            if (_dte != null)
            {
                _debuggerEvents = _dte.Events.DebuggerEvents;
                _debuggerEvents.OnEnterBreakMode += DebuggerEvents_OnEnterBreakMode;
                _debuggerEvents.OnContextChanged += DebuggerEvents_OnContextChanged;
                _debuggerEvents.OnEnterRunMode += DebuggerEvents_OnEnterRunMode;
                _debuggerEvents.OnEnterDesignMode += DebuggerEvents_OnEnterDesignMode;
            }

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

                return _callStackControl;
            }

            return null;
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

        #endregion
    }
}
