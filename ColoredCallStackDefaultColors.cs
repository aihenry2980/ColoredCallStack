using System.Drawing;

namespace ColorCallStack
{
    internal static class ColoredCallStackDefaultColors
    {
        internal static readonly Color LightNamespace = Color.FromArgb(255, 184, 182, 183);
        internal static readonly Color LightFunction = Color.FromArgb(255, 11, 94, 212);
        internal static readonly Color LightParamName = Color.FromArgb(255, 10, 99, 19);
        internal static readonly Color LightParamValue = Color.FromArgb(255, 201, 8, 198);
        internal static readonly Color LightFile = Color.FromArgb(255, 28, 98, 139);
        internal static readonly Color LightLine = Color.FromArgb(255, 150, 71, 11);
        internal static readonly Color LightPunctuation = Color.FromArgb(255, 33, 37, 41);

        internal static readonly Color DarkNamespace = Color.FromArgb(255, 156, 163, 175);
        internal static readonly Color DarkFunction = Color.FromArgb(255, 86, 156, 214);
        internal static readonly Color DarkParamName = Color.FromArgb(255, 206, 145, 120);
        internal static readonly Color DarkParamValue = Color.FromArgb(255, 181, 206, 168);
        internal static readonly Color DarkFile = Color.FromArgb(255, 156, 220, 254);
        internal static readonly Color DarkLine = Color.FromArgb(255, 255, 198, 109);
        internal static readonly Color DarkPunctuation = Color.FromArgb(255, 208, 208, 208);
    }
}
