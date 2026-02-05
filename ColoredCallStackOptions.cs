using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel;
using System.Drawing;

namespace ColorCallStack
{
    internal enum CallStackThemeMode
    {
        Auto = 0,
        Light = 1,
        Dark = 2
    }

    internal sealed class ColoredCallStackOptions : DialogPage
    {
        private CallStackThemeMode _themeMode = CallStackThemeMode.Auto;
        private bool _showParameterTypes;
        private Color _lightNamespaceColor = Color.FromArgb(255, 90, 96, 104);
        private Color _lightFunctionColor = Color.FromArgb(255, 0, 82, 153);
        private Color _lightParamNameColor = Color.FromArgb(255, 122, 63, 0);
        private Color _lightParamValueColor = Color.FromArgb(255, 0, 100, 0);
        private Color _lightFileColor = Color.FromArgb(255, 28, 98, 139);
        private Color _lightLineColor = Color.FromArgb(255, 170, 75, 0);
        private Color _lightPunctuationColor = Color.FromArgb(255, 33, 37, 41);
        private Color _darkNamespaceColor = Color.FromArgb(255, 156, 163, 175);
        private Color _darkFunctionColor = Color.FromArgb(255, 86, 156, 214);
        private Color _darkParamNameColor = Color.FromArgb(255, 206, 145, 120);
        private Color _darkParamValueColor = Color.FromArgb(255, 181, 206, 168);
        private Color _darkFileColor = Color.FromArgb(255, 156, 220, 254);
        private Color _darkLineColor = Color.FromArgb(255, 255, 198, 109);
        private Color _darkPunctuationColor = Color.FromArgb(255, 208, 208, 208);

        internal static event EventHandler ThemeModeChanged;

        [Category("Appearance")]
        [DisplayName("Theme Mode")]
        [Description("Choose whether the colored call stack follows the VS theme or forces light/dark.")]
        [DefaultValue(CallStackThemeMode.Auto)]
        public CallStackThemeMode ThemeMode
        {
            get => _themeMode;
            set => _themeMode = value;
        }

        [Category("Display")]
        [DisplayName("Show Parameter Types")]
        [Description("Show parameter types in the call stack.")]
        [DefaultValue(false)]
        public bool ShowParameterTypes
        {
            get => _showParameterTypes;
            set => _showParameterTypes = value;
        }

        [Category("Light Theme")]
        [DisplayName("Namespace")]
        [Description("Color for namespace/class prefixes in light theme.")]
        public Color LightNamespaceColor
        {
            get => _lightNamespaceColor;
            set => _lightNamespaceColor = value;
        }

        [Category("Light Theme")]
        [DisplayName("Function")]
        [Description("Color for function names in light theme.")]
        public Color LightFunctionColor
        {
            get => _lightFunctionColor;
            set => _lightFunctionColor = value;
        }

        [Category("Light Theme")]
        [DisplayName("Parameter Name")]
        [Description("Color for parameter names in light theme.")]
        public Color LightParamNameColor
        {
            get => _lightParamNameColor;
            set => _lightParamNameColor = value;
        }

        [Category("Light Theme")]
        [DisplayName("Parameter Value")]
        [Description("Color for parameter values in light theme.")]
        public Color LightParamValueColor
        {
            get => _lightParamValueColor;
            set => _lightParamValueColor = value;
        }

        [Category("Light Theme")]
        [DisplayName("File Name")]
        [Description("Color for file names in light theme.")]
        public Color LightFileColor
        {
            get => _lightFileColor;
            set => _lightFileColor = value;
        }

        [Category("Light Theme")]
        [DisplayName("Line Number")]
        [Description("Color for line numbers in light theme.")]
        public Color LightLineColor
        {
            get => _lightLineColor;
            set => _lightLineColor = value;
        }

        [Category("Light Theme")]
        [DisplayName("Punctuation")]
        [Description("Color for punctuation and separators in light theme.")]
        public Color LightPunctuationColor
        {
            get => _lightPunctuationColor;
            set => _lightPunctuationColor = value;
        }

        [Category("Dark Theme")]
        [DisplayName("Namespace")]
        [Description("Color for namespace/class prefixes in dark theme.")]
        public Color DarkNamespaceColor
        {
            get => _darkNamespaceColor;
            set => _darkNamespaceColor = value;
        }

        [Category("Dark Theme")]
        [DisplayName("Function")]
        [Description("Color for function names in dark theme.")]
        public Color DarkFunctionColor
        {
            get => _darkFunctionColor;
            set => _darkFunctionColor = value;
        }

        [Category("Dark Theme")]
        [DisplayName("Parameter Name")]
        [Description("Color for parameter names in dark theme.")]
        public Color DarkParamNameColor
        {
            get => _darkParamNameColor;
            set => _darkParamNameColor = value;
        }

        [Category("Dark Theme")]
        [DisplayName("Parameter Value")]
        [Description("Color for parameter values in dark theme.")]
        public Color DarkParamValueColor
        {
            get => _darkParamValueColor;
            set => _darkParamValueColor = value;
        }

        [Category("Dark Theme")]
        [DisplayName("File Name")]
        [Description("Color for file names in dark theme.")]
        public Color DarkFileColor
        {
            get => _darkFileColor;
            set => _darkFileColor = value;
        }

        [Category("Dark Theme")]
        [DisplayName("Line Number")]
        [Description("Color for line numbers in dark theme.")]
        public Color DarkLineColor
        {
            get => _darkLineColor;
            set => _darkLineColor = value;
        }

        [Category("Dark Theme")]
        [DisplayName("Punctuation")]
        [Description("Color for punctuation and separators in dark theme.")]
        public Color DarkPunctuationColor
        {
            get => _darkPunctuationColor;
            set => _darkPunctuationColor = value;
        }

        protected override void OnApply(PageApplyEventArgs e)
        {
            base.OnApply(e);
            ThemeModeChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
