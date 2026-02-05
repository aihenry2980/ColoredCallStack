using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;

namespace ColorCallStack
{
    /// <summary>
    /// This class implements the tool window exposed by this package and hosts a user control.
    /// </summary>
    /// <remarks>
    /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane,
    /// usually implemented by the package implementer.
    /// <para>
    /// This class derives from the ToolWindowPane class provided from the MPF in order to use its
    /// implementation of the IVsUIElementPane interface.
    /// </para>
    /// </remarks>
    [Guid("2e1a6457-746e-4e89-85e4-cdbc3ea6c35b")]
    public class ColoredCallStack : ToolWindowPane
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ColoredCallStack"/> class.
        /// </summary>
        public ColoredCallStack() : base(null)
        {
            this.Caption = "ColoredCallStack";

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
            // the object returned by the Content property.
            this.Content = new ColoredCallStackControl();
        }
    }
}
