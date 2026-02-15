using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Design;

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
        private bool _showNamespace = true;
        private bool _showLineNumbers = true;
        private bool _showFilePath = true;
        private int _fontSizeAdjustmentSteps;
        private int _namespaceFontSizePercent = 100;
        private int _functionFontSizePercent = 100;
        private int _parameterFontSizePercent = 100;
        private int _lineFontSizePercent = 100;
        private int _fileFontSizePercent = 100;
        private string _namespaceFontFace = string.Empty;
        private string _functionFontFace = string.Empty;
        private string _parameterFontFace = string.Empty;
        private string _lineFontFace = string.Empty;
        private string _fileFontFace = string.Empty;
        private Color _lightNamespaceColor = ColoredCallStackDefaultColors.LightNamespace;
        private Color _lightFunctionColor = ColoredCallStackDefaultColors.LightFunction;
        private Color _lightParamNameColor = ColoredCallStackDefaultColors.LightParamName;
        private Color _lightParamValueColor = ColoredCallStackDefaultColors.LightParamValue;
        private Color _lightFileColor = ColoredCallStackDefaultColors.LightFile;
        private Color _lightLineColor = ColoredCallStackDefaultColors.LightLine;
        private Color _lightPunctuationColor = ColoredCallStackDefaultColors.LightPunctuation;
        private Color _darkNamespaceColor = ColoredCallStackDefaultColors.DarkNamespace;
        private Color _darkFunctionColor = ColoredCallStackDefaultColors.DarkFunction;
        private Color _darkParamNameColor = ColoredCallStackDefaultColors.DarkParamName;
        private Color _darkParamValueColor = ColoredCallStackDefaultColors.DarkParamValue;
        private Color _darkFileColor = ColoredCallStackDefaultColors.DarkFile;
        private Color _darkLineColor = ColoredCallStackDefaultColors.DarkLine;
        private Color _darkPunctuationColor = ColoredCallStackDefaultColors.DarkPunctuation;

        private int _lightNamespaceColorArgb = ColoredCallStackDefaultColors.LightNamespace.ToArgb();
        private int _lightFunctionColorArgb = ColoredCallStackDefaultColors.LightFunction.ToArgb();
        private int _lightParamNameColorArgb = ColoredCallStackDefaultColors.LightParamName.ToArgb();
        private int _lightParamValueColorArgb = ColoredCallStackDefaultColors.LightParamValue.ToArgb();
        private int _lightFileColorArgb = ColoredCallStackDefaultColors.LightFile.ToArgb();
        private int _lightLineColorArgb = ColoredCallStackDefaultColors.LightLine.ToArgb();
        private int _lightPunctuationColorArgb = ColoredCallStackDefaultColors.LightPunctuation.ToArgb();
        private int _darkNamespaceColorArgb = ColoredCallStackDefaultColors.DarkNamespace.ToArgb();
        private int _darkFunctionColorArgb = ColoredCallStackDefaultColors.DarkFunction.ToArgb();
        private int _darkParamNameColorArgb = ColoredCallStackDefaultColors.DarkParamName.ToArgb();
        private int _darkParamValueColorArgb = ColoredCallStackDefaultColors.DarkParamValue.ToArgb();
        private int _darkFileColorArgb = ColoredCallStackDefaultColors.DarkFile.ToArgb();
        private int _darkLineColorArgb = ColoredCallStackDefaultColors.DarkLine.ToArgb();
        private int _darkPunctuationColorArgb = ColoredCallStackDefaultColors.DarkPunctuation.ToArgb();

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
        [DisplayName("Show Namespace")]
        [Description("Show namespace/class prefixes before function names.")]
        [DefaultValue(true)]
        public bool ShowNamespace
        {
            get => _showNamespace;
            set => _showNamespace = value;
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

        [Category("Display")]
        [DisplayName("Show Line Numbers")]
        [Description("Show line numbers in the call stack when available.")]
        [DefaultValue(true)]
        public bool ShowLineNumbers
        {
            get => _showLineNumbers;
            set => _showLineNumbers = value;
        }

        [Category("Display")]
        [DisplayName("Show File Path")]
        [Description("Show full file path instead of file name in the call stack when available.")]
        [DefaultValue(true)]
        public bool ShowFilePath
        {
            get => _showFilePath;
            set => _showFilePath = value;
        }

        [Browsable(false)]
        [DefaultValue(0)]
        public int FontSizeAdjustmentSteps
        {
            get => _fontSizeAdjustmentSteps;
            set => _fontSizeAdjustmentSteps = value;
        }

        [Category("Font")]
        [DisplayName("Namespace Size (%)")]
        [Description("Relative namespace font size percentage (50-300).")]
        [DefaultValue(100)]
        public int NamespaceFontSizePercent
        {
            get => _namespaceFontSizePercent;
            set => _namespaceFontSizePercent = ClampPercent(value);
        }

        [Category("Font")]
        [DisplayName("Function Size (%)")]
        [Description("Relative function font size percentage (50-300).")]
        [DefaultValue(100)]
        public int FunctionFontSizePercent
        {
            get => _functionFontSizePercent;
            set => _functionFontSizePercent = ClampPercent(value);
        }

        [Category("Font")]
        [DisplayName("Parameter Size (%)")]
        [Description("Relative parameter font size percentage (50-300).")]
        [DefaultValue(100)]
        public int ParameterFontSizePercent
        {
            get => _parameterFontSizePercent;
            set => _parameterFontSizePercent = ClampPercent(value);
        }

        [Category("Font")]
        [DisplayName("Line Number Size (%)")]
        [Description("Relative line number font size percentage (50-300).")]
        [DefaultValue(100)]
        public int LineFontSizePercent
        {
            get => _lineFontSizePercent;
            set => _lineFontSizePercent = ClampPercent(value);
        }

        [Category("Font")]
        [DisplayName("File Name Size (%)")]
        [Description("Relative file name font size percentage (50-300).")]
        [DefaultValue(100)]
        public int FileFontSizePercent
        {
            get => _fileFontSizePercent;
            set => _fileFontSizePercent = ClampPercent(value);
        }

        [Category("Font Face")]
        [DisplayName("Namespace Font Face")]
        [Description("Optional namespace/class font face. Leave empty to use Text Editor font.")]
        [DefaultValue("")]
        [TypeConverter(typeof(FontConverter.FontNameConverter))]
        [Editor(typeof(FontFaceEditor), typeof(UITypeEditor))]
        public string NamespaceFontFace
        {
            get => _namespaceFontFace;
            set => _namespaceFontFace = NormalizeFontFace(value);
        }

        [Category("Font Face")]
        [DisplayName("Function Font Face")]
        [Description("Optional function font face. Leave empty to use Text Editor font.")]
        [DefaultValue("")]
        [TypeConverter(typeof(FontConverter.FontNameConverter))]
        [Editor(typeof(FontFaceEditor), typeof(UITypeEditor))]
        public string FunctionFontFace
        {
            get => _functionFontFace;
            set => _functionFontFace = NormalizeFontFace(value);
        }

        [Category("Font Face")]
        [DisplayName("Parameter Font Face")]
        [Description("Optional parameter font face. Leave empty to use Text Editor font.")]
        [DefaultValue("")]
        [TypeConverter(typeof(FontConverter.FontNameConverter))]
        [Editor(typeof(FontFaceEditor), typeof(UITypeEditor))]
        public string ParameterFontFace
        {
            get => _parameterFontFace;
            set => _parameterFontFace = NormalizeFontFace(value);
        }

        [Category("Font Face")]
        [DisplayName("Line Number Font Face")]
        [Description("Optional line number font face. Leave empty to use Text Editor font.")]
        [DefaultValue("")]
        [TypeConverter(typeof(FontConverter.FontNameConverter))]
        [Editor(typeof(FontFaceEditor), typeof(UITypeEditor))]
        public string LineFontFace
        {
            get => _lineFontFace;
            set => _lineFontFace = NormalizeFontFace(value);
        }

        [Category("Font Face")]
        [DisplayName("File Name Font Face")]
        [Description("Optional file name font face. Leave empty to use Text Editor font.")]
        [DefaultValue("")]
        [TypeConverter(typeof(FontConverter.FontNameConverter))]
        [Editor(typeof(FontFaceEditor), typeof(UITypeEditor))]
        public string FileFontFace
        {
            get => _fileFontFace;
            set => _fileFontFace = NormalizeFontFace(value);
        }

        [Category("Light Theme")]
        [DisplayName("Namespace")]
        [Description("Color for namespace/class prefixes in light theme.")]
        public Color LightNamespaceColor
        {
            get => _lightNamespaceColor;
            set => SetColor(ref _lightNamespaceColor, ref _lightNamespaceColorArgb, value, ColoredCallStackDefaultColors.LightNamespace);
        }

        [Category("Light Theme")]
        [DisplayName("Function")]
        [Description("Color for function names in light theme.")]
        public Color LightFunctionColor
        {
            get => _lightFunctionColor;
            set => SetColor(ref _lightFunctionColor, ref _lightFunctionColorArgb, value, ColoredCallStackDefaultColors.LightFunction);
        }

        [Category("Light Theme")]
        [DisplayName("Parameter Name")]
        [Description("Color for parameter names in light theme.")]
        public Color LightParamNameColor
        {
            get => _lightParamNameColor;
            set => SetColor(ref _lightParamNameColor, ref _lightParamNameColorArgb, value, ColoredCallStackDefaultColors.LightParamName);
        }

        [Category("Light Theme")]
        [DisplayName("Parameter Value")]
        [Description("Color for parameter values in light theme.")]
        public Color LightParamValueColor
        {
            get => _lightParamValueColor;
            set => SetColor(ref _lightParamValueColor, ref _lightParamValueColorArgb, value, ColoredCallStackDefaultColors.LightParamValue);
        }

        [Category("Light Theme")]
        [DisplayName("File Name")]
        [Description("Color for file names in light theme.")]
        public Color LightFileColor
        {
            get => _lightFileColor;
            set => SetColor(ref _lightFileColor, ref _lightFileColorArgb, value, ColoredCallStackDefaultColors.LightFile);
        }

        [Category("Light Theme")]
        [DisplayName("Line Number")]
        [Description("Color for line numbers in light theme.")]
        public Color LightLineColor
        {
            get => _lightLineColor;
            set => SetColor(ref _lightLineColor, ref _lightLineColorArgb, value, ColoredCallStackDefaultColors.LightLine);
        }

        [Category("Light Theme")]
        [DisplayName("Punctuation")]
        [Description("Color for punctuation and separators in light theme.")]
        public Color LightPunctuationColor
        {
            get => _lightPunctuationColor;
            set => SetColor(ref _lightPunctuationColor, ref _lightPunctuationColorArgb, value, ColoredCallStackDefaultColors.LightPunctuation);
        }

        [Category("Dark Theme")]
        [DisplayName("Namespace")]
        [Description("Color for namespace/class prefixes in dark theme.")]
        public Color DarkNamespaceColor
        {
            get => _darkNamespaceColor;
            set => SetColor(ref _darkNamespaceColor, ref _darkNamespaceColorArgb, value, ColoredCallStackDefaultColors.DarkNamespace);
        }

        [Category("Dark Theme")]
        [DisplayName("Function")]
        [Description("Color for function names in dark theme.")]
        public Color DarkFunctionColor
        {
            get => _darkFunctionColor;
            set => SetColor(ref _darkFunctionColor, ref _darkFunctionColorArgb, value, ColoredCallStackDefaultColors.DarkFunction);
        }

        [Category("Dark Theme")]
        [DisplayName("Parameter Name")]
        [Description("Color for parameter names in dark theme.")]
        public Color DarkParamNameColor
        {
            get => _darkParamNameColor;
            set => SetColor(ref _darkParamNameColor, ref _darkParamNameColorArgb, value, ColoredCallStackDefaultColors.DarkParamName);
        }

        [Category("Dark Theme")]
        [DisplayName("Parameter Value")]
        [Description("Color for parameter values in dark theme.")]
        public Color DarkParamValueColor
        {
            get => _darkParamValueColor;
            set => SetColor(ref _darkParamValueColor, ref _darkParamValueColorArgb, value, ColoredCallStackDefaultColors.DarkParamValue);
        }

        [Category("Dark Theme")]
        [DisplayName("File Name")]
        [Description("Color for file names in dark theme.")]
        public Color DarkFileColor
        {
            get => _darkFileColor;
            set => SetColor(ref _darkFileColor, ref _darkFileColorArgb, value, ColoredCallStackDefaultColors.DarkFile);
        }

        [Category("Dark Theme")]
        [DisplayName("Line Number")]
        [Description("Color for line numbers in dark theme.")]
        public Color DarkLineColor
        {
            get => _darkLineColor;
            set => SetColor(ref _darkLineColor, ref _darkLineColorArgb, value, ColoredCallStackDefaultColors.DarkLine);
        }

        [Category("Dark Theme")]
        [DisplayName("Punctuation")]
        [Description("Color for punctuation and separators in dark theme.")]
        public Color DarkPunctuationColor
        {
            get => _darkPunctuationColor;
            set => SetColor(ref _darkPunctuationColor, ref _darkPunctuationColorArgb, value, ColoredCallStackDefaultColors.DarkPunctuation);
        }

        protected override void OnApply(PageApplyEventArgs e)
        {
            base.OnApply(e);
            SyncColorArgbShadowValues();
            SaveSettingsToStorage();
            ThemeModeChanged?.Invoke(this, EventArgs.Empty);
        }

        internal void ResetToFactoryDefaults()
        {
            ThemeMode = CallStackThemeMode.Auto;
            ShowNamespace = true;
            ShowParameterTypes = false;
            ShowLineNumbers = true;
            ShowFilePath = true;
            FontSizeAdjustmentSteps = 0;

            NamespaceFontSizePercent = 100;
            FunctionFontSizePercent = 100;
            ParameterFontSizePercent = 100;
            LineFontSizePercent = 100;
            FileFontSizePercent = 100;

            NamespaceFontFace = string.Empty;
            FunctionFontFace = string.Empty;
            ParameterFontFace = string.Empty;
            LineFontFace = string.Empty;
            FileFontFace = string.Empty;

            LightNamespaceColor = ColoredCallStackDefaultColors.LightNamespace;
            LightFunctionColor = ColoredCallStackDefaultColors.LightFunction;
            LightParamNameColor = ColoredCallStackDefaultColors.LightParamName;
            LightParamValueColor = ColoredCallStackDefaultColors.LightParamValue;
            LightFileColor = ColoredCallStackDefaultColors.LightFile;
            LightLineColor = ColoredCallStackDefaultColors.LightLine;
            LightPunctuationColor = ColoredCallStackDefaultColors.LightPunctuation;
            DarkNamespaceColor = ColoredCallStackDefaultColors.DarkNamespace;
            DarkFunctionColor = ColoredCallStackDefaultColors.DarkFunction;
            DarkParamNameColor = ColoredCallStackDefaultColors.DarkParamName;
            DarkParamValueColor = ColoredCallStackDefaultColors.DarkParamValue;
            DarkFileColor = ColoredCallStackDefaultColors.DarkFile;
            DarkLineColor = ColoredCallStackDefaultColors.DarkLine;
            DarkPunctuationColor = ColoredCallStackDefaultColors.DarkPunctuation;
        }

        public override void LoadSettingsFromStorage()
        {
            base.LoadSettingsFromStorage();
            RestoreColorsFromArgbFallbacks();
        }

        public override void SaveSettingsToStorage()
        {
            SyncColorArgbShadowValues();
            base.SaveSettingsToStorage();
        }

        [Browsable(false)]
        public int LightNamespaceColorArgb
        {
            get => _lightNamespaceColorArgb;
            set => SetColorFromArgb(ref _lightNamespaceColor, ref _lightNamespaceColorArgb, value, ColoredCallStackDefaultColors.LightNamespace);
        }

        [Browsable(false)]
        public int LightFunctionColorArgb
        {
            get => _lightFunctionColorArgb;
            set => SetColorFromArgb(ref _lightFunctionColor, ref _lightFunctionColorArgb, value, ColoredCallStackDefaultColors.LightFunction);
        }

        [Browsable(false)]
        public int LightParamNameColorArgb
        {
            get => _lightParamNameColorArgb;
            set => SetColorFromArgb(ref _lightParamNameColor, ref _lightParamNameColorArgb, value, ColoredCallStackDefaultColors.LightParamName);
        }

        [Browsable(false)]
        public int LightParamValueColorArgb
        {
            get => _lightParamValueColorArgb;
            set => SetColorFromArgb(ref _lightParamValueColor, ref _lightParamValueColorArgb, value, ColoredCallStackDefaultColors.LightParamValue);
        }

        [Browsable(false)]
        public int LightFileColorArgb
        {
            get => _lightFileColorArgb;
            set => SetColorFromArgb(ref _lightFileColor, ref _lightFileColorArgb, value, ColoredCallStackDefaultColors.LightFile);
        }

        [Browsable(false)]
        public int LightLineColorArgb
        {
            get => _lightLineColorArgb;
            set => SetColorFromArgb(ref _lightLineColor, ref _lightLineColorArgb, value, ColoredCallStackDefaultColors.LightLine);
        }

        [Browsable(false)]
        public int LightPunctuationColorArgb
        {
            get => _lightPunctuationColorArgb;
            set => SetColorFromArgb(ref _lightPunctuationColor, ref _lightPunctuationColorArgb, value, ColoredCallStackDefaultColors.LightPunctuation);
        }

        [Browsable(false)]
        public int DarkNamespaceColorArgb
        {
            get => _darkNamespaceColorArgb;
            set => SetColorFromArgb(ref _darkNamespaceColor, ref _darkNamespaceColorArgb, value, ColoredCallStackDefaultColors.DarkNamespace);
        }

        [Browsable(false)]
        public int DarkFunctionColorArgb
        {
            get => _darkFunctionColorArgb;
            set => SetColorFromArgb(ref _darkFunctionColor, ref _darkFunctionColorArgb, value, ColoredCallStackDefaultColors.DarkFunction);
        }

        [Browsable(false)]
        public int DarkParamNameColorArgb
        {
            get => _darkParamNameColorArgb;
            set => SetColorFromArgb(ref _darkParamNameColor, ref _darkParamNameColorArgb, value, ColoredCallStackDefaultColors.DarkParamName);
        }

        [Browsable(false)]
        public int DarkParamValueColorArgb
        {
            get => _darkParamValueColorArgb;
            set => SetColorFromArgb(ref _darkParamValueColor, ref _darkParamValueColorArgb, value, ColoredCallStackDefaultColors.DarkParamValue);
        }

        [Browsable(false)]
        public int DarkFileColorArgb
        {
            get => _darkFileColorArgb;
            set => SetColorFromArgb(ref _darkFileColor, ref _darkFileColorArgb, value, ColoredCallStackDefaultColors.DarkFile);
        }

        [Browsable(false)]
        public int DarkLineColorArgb
        {
            get => _darkLineColorArgb;
            set => SetColorFromArgb(ref _darkLineColor, ref _darkLineColorArgb, value, ColoredCallStackDefaultColors.DarkLine);
        }

        [Browsable(false)]
        public int DarkPunctuationColorArgb
        {
            get => _darkPunctuationColorArgb;
            set => SetColorFromArgb(ref _darkPunctuationColor, ref _darkPunctuationColorArgb, value, ColoredCallStackDefaultColors.DarkPunctuation);
        }

        private static int ClampPercent(int value)
        {
            if (value < 50)
            {
                return 50;
            }

            if (value > 300)
            {
                return 300;
            }

            return value;
        }

        private static string NormalizeFontFace(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        private static void SetColor(ref Color colorField, ref int argbField, Color value, Color fallback)
        {
            if (value.IsEmpty)
            {
                colorField = fallback;
                argbField = fallback.ToArgb();
                return;
            }

            colorField = value;
            argbField = value.ToArgb();
        }

        private static void SetColorFromArgb(ref Color colorField, ref int argbField, int value, Color fallback)
        {
            argbField = value;
            colorField = value == 0 ? fallback : Color.FromArgb(value);
        }

        private void SyncColorArgbShadowValues()
        {
            _lightNamespaceColorArgb = ResolveNonEmptyColor(_lightNamespaceColor, ColoredCallStackDefaultColors.LightNamespace).ToArgb();
            _lightFunctionColorArgb = ResolveNonEmptyColor(_lightFunctionColor, ColoredCallStackDefaultColors.LightFunction).ToArgb();
            _lightParamNameColorArgb = ResolveNonEmptyColor(_lightParamNameColor, ColoredCallStackDefaultColors.LightParamName).ToArgb();
            _lightParamValueColorArgb = ResolveNonEmptyColor(_lightParamValueColor, ColoredCallStackDefaultColors.LightParamValue).ToArgb();
            _lightFileColorArgb = ResolveNonEmptyColor(_lightFileColor, ColoredCallStackDefaultColors.LightFile).ToArgb();
            _lightLineColorArgb = ResolveNonEmptyColor(_lightLineColor, ColoredCallStackDefaultColors.LightLine).ToArgb();
            _lightPunctuationColorArgb = ResolveNonEmptyColor(_lightPunctuationColor, ColoredCallStackDefaultColors.LightPunctuation).ToArgb();
            _darkNamespaceColorArgb = ResolveNonEmptyColor(_darkNamespaceColor, ColoredCallStackDefaultColors.DarkNamespace).ToArgb();
            _darkFunctionColorArgb = ResolveNonEmptyColor(_darkFunctionColor, ColoredCallStackDefaultColors.DarkFunction).ToArgb();
            _darkParamNameColorArgb = ResolveNonEmptyColor(_darkParamNameColor, ColoredCallStackDefaultColors.DarkParamName).ToArgb();
            _darkParamValueColorArgb = ResolveNonEmptyColor(_darkParamValueColor, ColoredCallStackDefaultColors.DarkParamValue).ToArgb();
            _darkFileColorArgb = ResolveNonEmptyColor(_darkFileColor, ColoredCallStackDefaultColors.DarkFile).ToArgb();
            _darkLineColorArgb = ResolveNonEmptyColor(_darkLineColor, ColoredCallStackDefaultColors.DarkLine).ToArgb();
            _darkPunctuationColorArgb = ResolveNonEmptyColor(_darkPunctuationColor, ColoredCallStackDefaultColors.DarkPunctuation).ToArgb();
        }

        private void RestoreColorsFromArgbFallbacks()
        {
            _lightNamespaceColor = ResolveLoadedColor(_lightNamespaceColor, _lightNamespaceColorArgb, ColoredCallStackDefaultColors.LightNamespace);
            _lightFunctionColor = ResolveLoadedColor(_lightFunctionColor, _lightFunctionColorArgb, ColoredCallStackDefaultColors.LightFunction);
            _lightParamNameColor = ResolveLoadedColor(_lightParamNameColor, _lightParamNameColorArgb, ColoredCallStackDefaultColors.LightParamName);
            _lightParamValueColor = ResolveLoadedColor(_lightParamValueColor, _lightParamValueColorArgb, ColoredCallStackDefaultColors.LightParamValue);
            _lightFileColor = ResolveLoadedColor(_lightFileColor, _lightFileColorArgb, ColoredCallStackDefaultColors.LightFile);
            _lightLineColor = ResolveLoadedColor(_lightLineColor, _lightLineColorArgb, ColoredCallStackDefaultColors.LightLine);
            _lightPunctuationColor = ResolveLoadedColor(_lightPunctuationColor, _lightPunctuationColorArgb, ColoredCallStackDefaultColors.LightPunctuation);
            _darkNamespaceColor = ResolveLoadedColor(_darkNamespaceColor, _darkNamespaceColorArgb, ColoredCallStackDefaultColors.DarkNamespace);
            _darkFunctionColor = ResolveLoadedColor(_darkFunctionColor, _darkFunctionColorArgb, ColoredCallStackDefaultColors.DarkFunction);
            _darkParamNameColor = ResolveLoadedColor(_darkParamNameColor, _darkParamNameColorArgb, ColoredCallStackDefaultColors.DarkParamName);
            _darkParamValueColor = ResolveLoadedColor(_darkParamValueColor, _darkParamValueColorArgb, ColoredCallStackDefaultColors.DarkParamValue);
            _darkFileColor = ResolveLoadedColor(_darkFileColor, _darkFileColorArgb, ColoredCallStackDefaultColors.DarkFile);
            _darkLineColor = ResolveLoadedColor(_darkLineColor, _darkLineColorArgb, ColoredCallStackDefaultColors.DarkLine);
            _darkPunctuationColor = ResolveLoadedColor(_darkPunctuationColor, _darkPunctuationColorArgb, ColoredCallStackDefaultColors.DarkPunctuation);
            SyncColorArgbShadowValues();
        }

        private static Color ResolveLoadedColor(Color loadedColor, int loadedArgb, Color fallback)
        {
            if (!loadedColor.IsEmpty)
            {
                return loadedColor;
            }

            if (loadedArgb != 0)
            {
                return Color.FromArgb(loadedArgb);
            }

            return fallback;
        }

        private static Color ResolveNonEmptyColor(Color color, Color fallback)
        {
            return color.IsEmpty ? fallback : color;
        }
    }
}

