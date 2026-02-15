using EnvDTE;
using EnvDTE90a;
using Microsoft.VisualStudio.Debugger.Interop;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Reflection;
using Task = System.Threading.Tasks.Task;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Text.RegularExpressions;

namespace ColorCallStack
{
    /// <summary>
    /// Interaction logic for ColoredCallStackControl.
    /// </summary>
    public partial class ColoredCallStackControl : UserControl
    {
        private Brush _currentFrameBrush;
        private Brush _namespaceBrush;
        private Brush _functionBrush;
        private Brush _paramNameBrush;
        private Brush _paramValueBrush;
        private Brush _fileBrush;
        private Brush _lineBrush;
        private Brush _punctuationBrush;
        private FontFamily _fontFamily = new FontFamily("Consolas");
        private double _fontSize = 12.0;
        private Palette _lightPalette = LightPalette;
        private Palette _darkPalette = DarkPalette;
        private bool _showNamespace = true;
        private bool _showParameterTypes;
        private bool _showLineNumbers = true;
        private bool _showFilePath;
        private bool _hexDisplayMode = true;
        private double _namespaceFontScale = 1.0;
        private double _functionFontScale = 1.0;
        private double _parameterFontScale = 1.0;
        private double _lineFontScale = 1.0;
        private double _fileFontScale = 1.0;
        private string _namespaceFontFamilyName;
        private string _functionFontFamilyName;
        private string _parameterFontFamilyName;
        private string _lineFontFamilyName;
        private string _fileFontFamilyName;
        private FontFamily _namespaceFontFamily;
        private FontFamily _functionFontFamily;
        private FontFamily _parameterFontFamily;
        private FontFamily _lineFontFamily;
        private FontFamily _fileFontFamily;
        private CancellationTokenSource _detailsCts;
        private bool _themeHooked;
        private List<StackFrame> _lastFrames;
        private StackFrame _lastCurrentFrame;
        private CallStackThemeMode _themeMode = CallStackThemeMode.Auto;
        private bool _skipNextDetailsDelay;

        internal event EventHandler<StackFrame> FrameActivated;
        internal event EventHandler<DisplayOptionsChangedEventArgs> DisplayOptionsChanged;
        internal event EventHandler<int> FontSizeStepRequested;
        internal event EventHandler ResetRequested;

        /// <summary>
        /// Initializes a new instance of the <see cref="ColoredCallStackControl"/> class.
        /// </summary>
        public ColoredCallStackControl()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            InitializeComponent();
            UpdateThemeResources();
            Loaded += OnLoaded;
            Unloaded += OnUnloaded;
        }

        internal void SetThemeMode(CallStackThemeMode mode)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_themeMode == mode)
            {
                return;
            }

            _themeMode = mode;
            UpdateThemeResources();
            if (_lastFrames != null)
            {
                UpdateCallStack(_lastFrames, _lastCurrentFrame);
            }
        }

        internal void SetPalettes(Palette lightPalette, Palette darkPalette)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _lightPalette = lightPalette;
            _darkPalette = darkPalette;
            UpdateThemeResources();
            if (_lastFrames != null)
            {
                UpdateCallStack(_lastFrames, _lastCurrentFrame);
            }
        }

        internal void SetDisplayOptions(bool showNamespace, bool showParameterTypes, bool showLineNumbers, bool showFilePath, bool hexDisplayMode)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            bool changed = _showNamespace != showNamespace ||
                           _showParameterTypes != showParameterTypes ||
                           _showLineNumbers != showLineNumbers ||
                           _showFilePath != showFilePath ||
                           _hexDisplayMode != hexDisplayMode;

            _showNamespace = showNamespace;
            _showParameterTypes = showParameterTypes;
            _showLineNumbers = showLineNumbers;
            _showFilePath = showFilePath;
            _hexDisplayMode = hexDisplayMode;
            UpdateContextMenuChecks();

            if (changed && _lastFrames != null)
            {
                UpdateCallStack(_lastFrames, _lastCurrentFrame);
            }
        }

        internal void SetTokenFontScales(double namespaceScale, double functionScale, double parameterScale)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            SetTokenFontScales(namespaceScale, functionScale, parameterScale, 1.0, 1.0);
        }

        internal void SetTokenFontScales(double namespaceScale, double functionScale, double parameterScale, double lineScale, double fileScale)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            namespaceScale = ClampScale(namespaceScale);
            functionScale = ClampScale(functionScale);
            parameterScale = ClampScale(parameterScale);
            lineScale = ClampScale(lineScale);
            fileScale = ClampScale(fileScale);

            if (Math.Abs(_namespaceFontScale - namespaceScale) < 0.001 &&
                Math.Abs(_functionFontScale - functionScale) < 0.001 &&
                Math.Abs(_parameterFontScale - parameterScale) < 0.001 &&
                Math.Abs(_lineFontScale - lineScale) < 0.001 &&
                Math.Abs(_fileFontScale - fileScale) < 0.001)
            {
                return;
            }

            _namespaceFontScale = namespaceScale;
            _functionFontScale = functionScale;
            _parameterFontScale = parameterScale;
            _lineFontScale = lineScale;
            _fileFontScale = fileScale;

            if (_lastFrames != null)
            {
                UpdateCallStack(_lastFrames, _lastCurrentFrame);
            }
        }

        internal void SetTokenFontFamilies(string namespaceFontFamily, string functionFontFamily, string parameterFontFamily, string lineFontFamily, string fileFontFamily)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string normalizedNamespace = NormalizeFontFamilyName(namespaceFontFamily);
            string normalizedFunction = NormalizeFontFamilyName(functionFontFamily);
            string normalizedParameter = NormalizeFontFamilyName(parameterFontFamily);
            string normalizedLine = NormalizeFontFamilyName(lineFontFamily);
            string normalizedFile = NormalizeFontFamilyName(fileFontFamily);

            if (string.Equals(_namespaceFontFamilyName, normalizedNamespace, StringComparison.Ordinal) &&
                string.Equals(_functionFontFamilyName, normalizedFunction, StringComparison.Ordinal) &&
                string.Equals(_parameterFontFamilyName, normalizedParameter, StringComparison.Ordinal) &&
                string.Equals(_lineFontFamilyName, normalizedLine, StringComparison.Ordinal) &&
                string.Equals(_fileFontFamilyName, normalizedFile, StringComparison.Ordinal))
            {
                return;
            }

            _namespaceFontFamilyName = normalizedNamespace;
            _functionFontFamilyName = normalizedFunction;
            _parameterFontFamilyName = normalizedParameter;
            _lineFontFamilyName = normalizedLine;
            _fileFontFamilyName = normalizedFile;

            _namespaceFontFamily = CreateFontFamilyOrNull(_namespaceFontFamilyName);
            _functionFontFamily = CreateFontFamilyOrNull(_functionFontFamilyName);
            _parameterFontFamily = CreateFontFamilyOrNull(_parameterFontFamilyName);
            _lineFontFamily = CreateFontFamilyOrNull(_lineFontFamilyName);
            _fileFontFamily = CreateFontFamilyOrNull(_fileFontFamilyName);

            if (_lastFrames != null)
            {
                UpdateCallStack(_lastFrames, _lastCurrentFrame);
            }
        }

        internal void SetFont(string fontFamilyName, double fontSize)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!string.IsNullOrWhiteSpace(fontFamilyName))
            {
                try
                {
                    _fontFamily = new FontFamily(fontFamilyName);
                }
                catch
                {
                    _fontFamily = new FontFamily("Consolas");
                }
            }

            if (fontSize > 0)
            {
                _fontSize = fontSize;
            }

            if (_lastFrames != null)
            {
                UpdateCallStack(_lastFrames, _lastCurrentFrame);
            }
        }

        public void ClearCallStack()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            FramesList.Items.Clear();
            CancelDetailsPopulation();
        }

        public void UpdateCallStack(IReadOnlyList<StackFrame> frames, StackFrame currentFrame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _lastFrames = frames != null ? new List<StackFrame>(frames) : null;
            _lastCurrentFrame = currentFrame;
            FramesList.Items.Clear();
            if (frames == null || frames.Count == 0)
            {
                CancelDetailsPopulation();
                return;
            }

            ListBoxItem currentItem = null;
            foreach (var frame in frames)
            {
                bool isCurrent = IsSameFrame(frame, currentFrame);
                ListBoxItem item;
                try
                {
                    item = CreateFrameItem(frame, isCurrent);
                }
                catch
                {
                    item = CreateFallbackItem(frame, isCurrent);
                }
                FramesList.Items.Add(item);
                if (isCurrent)
                {
                    currentItem = item;
                }
            }

            if (currentItem != null)
            {
                FramesList.SelectedItem = currentItem;
                FramesList.ScrollIntoView(currentItem);
            }

            BeginPopulateDetails();
        }

        private ListBoxItem CreateFrameItem(StackFrame frame, bool isCurrent)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var border = new Border
            {
                Background = isCurrent ? _currentFrameBrush : Brushes.Transparent,
                Padding = new Thickness(2, 1, 2, 1)
            };

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var arrow = new CrispImage
            {
                Width = 16,
                Height = 16,
                Margin = new Thickness(1, 0, 4, 0),
                Moniker = KnownMonikers.CurrentInstructionPointer,
                Visibility = isCurrent ? Visibility.Visible : Visibility.Hidden,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(arrow, 0);

            var text = new TextBlock
            {
                FontFamily = _fontFamily,
                FontSize = _fontSize,
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping = TextWrapping.NoWrap,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            BuildFrameInlines(text, frame, includeArguments: false);
            Grid.SetColumn(text, 1);

            var fileText = new TextBlock
            {
                FontFamily = _fontFamily,
                FontSize = _fontSize,
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping = TextWrapping.NoWrap,
                Margin = new Thickness(12, 0, 6, 0),
                HorizontalAlignment = HorizontalAlignment.Right
            };
            BuildFileInlines(fileText, frame, includeFileInfo: true);
            Grid.SetColumn(fileText, 2);

            grid.Children.Add(arrow);
            grid.Children.Add(text);
            grid.Children.Add(fileText);
            border.Child = grid;

            var item = new ListBoxItem
            {
                Content = border,
                Tag = new FrameRowInfo(frame, border, text, fileText, isCurrent)
            };

            return item;
        }

        private void BuildFrameInlines(TextBlock textBlock, StackFrame frame, bool includeArguments)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (textBlock == null)
            {
                return;
            }

            if (frame == null)
            {
                AddRun(textBlock, "<unknown>", _functionBrush);
                return;
            }

            string function = GetSafeFunctionName(frame);
            SplitFunctionName(function, out string namespacePart, out string functionPart);

            if (_showNamespace && !string.IsNullOrEmpty(namespacePart))
            {
                AddRun(textBlock, namespacePart, _namespaceBrush, _namespaceFontScale, _namespaceFontFamily);
            }

            AddRun(textBlock, functionPart, _functionBrush, _functionFontScale, _functionFontFamily);

            if (includeArguments && TryGetArguments(frame.Arguments, out List<ArgumentPart> args) && args.Count > 0)
            {
                AddRun(textBlock, "(", _punctuationBrush);
                for (int i = 0; i < args.Count; i++)
                {
                    if (i > 0)
                    {
                        AddRun(textBlock, ", ", _punctuationBrush);
                    }

                    ArgumentPart arg = args[i];
                    if (_showParameterTypes && !string.IsNullOrEmpty(arg.Type))
                    {
                        AddRun(textBlock, arg.Type, _paramNameBrush, _parameterFontScale, _parameterFontFamily);
                        AddRun(textBlock, " ", _punctuationBrush);
                    }

                    if (!string.IsNullOrEmpty(arg.Name))
                    {
                        AddRun(textBlock, arg.Name, _paramNameBrush, _parameterFontScale, _parameterFontFamily);
                        if (!string.IsNullOrEmpty(arg.Value))
                        {
                            AddRun(textBlock, "=", _punctuationBrush);
                            AddRun(textBlock, arg.Value, _paramValueBrush, _parameterFontScale, _parameterFontFamily);
                        }
                    }
                    else if (!string.IsNullOrEmpty(arg.Value))
                    {
                        AddRun(textBlock, arg.Value, _paramValueBrush, _parameterFontScale, _parameterFontFamily);
                    }
                }
                AddRun(textBlock, ")", _punctuationBrush);
            }

            AddInlineLineNumber(textBlock, frame);

        }

        private void AddInlineLineNumber(TextBlock textBlock, StackFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!_showLineNumbers || textBlock == null || frame == null)
            {
                return;
            }

            if (TryGetLineNumber(frame, out int line) && line > 0)
            {
                AddRun(textBlock, " Line ", _punctuationBrush, _lineFontScale, _lineFontFamily);
                AddRun(textBlock, line.ToString(), _lineBrush, _lineFontScale, _lineFontFamily);
            }
        }

        private void BuildFileInlines(TextBlock textBlock, StackFrame frame, bool includeFileInfo)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (textBlock == null)
            {
                return;
            }

            if (!includeFileInfo)
            {
                return;
            }

            if (TryGetFileInfo(frame, out string file, out int _))
            {
                if (!_showFilePath)
                {
                    return;
                }

                string fileText = System.IO.Path.GetFileName(file);
                if (string.IsNullOrEmpty(fileText))
                {
                    fileText = file;
                }

                AddRun(textBlock, fileText, _fileBrush, _fileFontScale, _fileFontFamily);
            }
        }

        private ListBoxItem CreateFallbackItem(StackFrame frame, bool isCurrent)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string function = GetSafeFunctionName(frame);
            var border = new Border
            {
                Background = isCurrent ? _currentFrameBrush : Brushes.Transparent,
                Padding = new Thickness(2, 1, 2, 1)
            };

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var arrow = new CrispImage
            {
                Width = 16,
                Height = 16,
                Margin = new Thickness(1, 0, 4, 0),
                Moniker = KnownMonikers.CurrentInstructionPointer,
                Visibility = isCurrent ? Visibility.Visible : Visibility.Hidden,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(arrow, 0);

            var text = new TextBlock
            {
                FontFamily = _fontFamily,
                FontSize = _fontSize,
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping = TextWrapping.NoWrap,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            string fallbackText = function;
            if (!_showNamespace)
            {
                SplitFunctionName(function, out _, out string functionPart);
                fallbackText = functionPart;
            }
            AddRun(text, fallbackText, _functionBrush, _functionFontScale, _functionFontFamily);
            AddInlineLineNumber(text, frame);
            Grid.SetColumn(text, 1);

            var fileText = new TextBlock
            {
                FontFamily = _fontFamily,
                FontSize = _fontSize,
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping = TextWrapping.NoWrap,
                Margin = new Thickness(12, 0, 6, 0),
                HorizontalAlignment = HorizontalAlignment.Right
            };
            BuildFileInlines(fileText, frame, includeFileInfo: true);
            Grid.SetColumn(fileText, 2);

            grid.Children.Add(arrow);
            grid.Children.Add(text);
            grid.Children.Add(fileText);
            border.Child = grid;

            return new ListBoxItem
            {
                Content = border,
                Tag = new FrameRowInfo(frame, border, text, fileText, isCurrent)
            };
        }

        private static string GetSafeFunctionName(StackFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (frame == null)
            {
                return "<unknown>";
            }

            try
            {
                return string.IsNullOrWhiteSpace(frame.FunctionName) ? "<unknown>" : frame.FunctionName;
            }
            catch
            {
                return "<unknown>";
            }
        }

        private static bool IsSameFrame(StackFrame left, StackFrame right)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (left == null || right == null)
            {
                return false;
            }

            if (ReferenceEquals(left, right))
            {
                return true;
            }

            try
            {
                string leftFunction = left.FunctionName ?? string.Empty;
                string rightFunction = right.FunctionName ?? string.Empty;
                if (!string.Equals(leftFunction, rightFunction, StringComparison.Ordinal))
                {
                    return false;
                }

                if (TryGetFileInfo(left, out string leftFile, out int leftLine) &&
                    TryGetFileInfo(right, out string rightFile, out int rightLine))
                {
                    return string.Equals(leftFile, rightFile, StringComparison.OrdinalIgnoreCase) &&
                           leftLine == rightLine;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetArguments(Expressions expressions, out List<ArgumentPart> arguments)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            arguments = null;
            if (expressions == null)
            {
                return false;
            }

            try
            {
                if (expressions.Count == 0)
                {
                    return false;
                }

                var parts = new List<ArgumentPart>(expressions.Count);
                for (int i = 1; i <= expressions.Count; i++)
                {
                    EnvDTE.Expression expr = expressions.Item(i);
                    if (expr == null)
                    {
                        continue;
                    }

                    string name = expr.Name;
                    string value = expr.Value;
                    string type = null;
                    try
                    {
                        type = expr.Type;
                    }
                    catch
                    {
                        type = null;
                    }

                    parts.Add(new ArgumentPart(name, value, type));
                }

                arguments = parts;
                return parts.Count > 0;
            }
            catch
            {
                arguments = null;
                return false;
            }
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

        private static bool TryGetFileInfo(StackFrame frame, out string file, out int line)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            file = null;
            line = 0;
            if (frame == null)
            {
                return false;
            }

            if (TryGetFileInfoFromStackFrame2(frame, out file, out line))
            {
                return true;
            }

            if (TryGetStringPropertyValue(frame, "FileName", out string candidateFile))
            {
                file = candidateFile;
            }

            if (TryGetIntPropertyValue(frame, "LineNumber", out int candidateLine))
            {
                line = candidateLine;
            }

            if (!string.IsNullOrEmpty(file))
            {
                return true;
            }

            if (TryGetFileInfoFromDebugFrame(frame, out file, out line))
            {
                return true;
            }
            return false;
        }

        private static bool TryGetFileInfoFromStackFrame2(StackFrame frame, out string file, out int line)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            file = null;
            line = 0;
            if (!(frame is StackFrame2 frame2))
            {
                return false;
            }

            try
            {
                file = frame2.FileName;
            }
            catch
            {
                file = null;
            }

            try
            {
                line = unchecked((int)frame2.LineNumber);
            }
            catch
            {
                line = 0;
            }

            return !string.IsNullOrEmpty(file);
        }

        private static bool TryGetLineNumber(StackFrame frame, out int line)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            line = 0;
            if (frame == null)
            {
                return false;
            }

            if (TryGetLineNumberFromStackFrame2(frame, out line))
            {
                return true;
            }

            if ((TryGetIntPropertyValue(frame, "LineNumber", out line) ||
                 TryGetIntPropertyValue(frame, "Line", out line)) &&
                line > 0)
            {
                return true;
            }

            if (TryGetLineNumberFromFrameText(frame, out line))
            {
                return true;
            }

            if (!TryGetDebugFrame(frame, out IDebugStackFrame2 debugFrame))
            {
                return false;
            }

            if (TryGetLineNumberFromFrameInfo(debugFrame, out line))
            {
                return true;
            }

            if (TryGetDocumentContext(debugFrame, out IDebugDocumentContext2 context) &&
                TryGetLineNumberFromDocumentContext(context, out line))
            {
                return true;
            }

            return false;
        }

        private static bool TryGetLineNumberFromStackFrame2(StackFrame frame, out int line)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            line = 0;
            if (!(frame is StackFrame2 frame2))
            {
                return false;
            }

            try
            {
                line = unchecked((int)frame2.LineNumber);
                return line > 0;
            }
            catch
            {
                line = 0;
                return false;
            }
        }

        private static bool TryGetLineNumberFromFrameText(StackFrame frame, out int line)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            line = 0;
            if (frame == null)
            {
                return false;
            }

            if (TryGetStringPropertyValue(frame, "Name", out string frameName) &&
                TryParseLineFromText(frameName, out line))
            {
                return true;
            }

            try
            {
                if (TryParseLineFromText(frame.FunctionName, out line))
                {
                    return true;
                }
            }
            catch
            {
                line = 0;
            }

            try
            {
                if (TryParseLineFromText(frame.ToString(), out line))
                {
                    return true;
                }
            }
            catch
            {
                line = 0;
            }

            return false;
        }

        private static bool TryGetFileInfoFromDebugFrame(StackFrame frame, out string file, out int line)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            file = null;
            line = 0;

            if (!TryGetDebugFrame(frame, out IDebugStackFrame2 debugFrame))
            {
                return false;
            }

            try
            {
                if (TryGetFileInfoFromFrameInfo(debugFrame, out file, out line))
                {
                    return true;
                }

                if (!TryGetDocumentContext(debugFrame, out IDebugDocumentContext2 context))
                {
                    return false;
                }

                return TryGetFileInfoFromDocumentContext(context, out file, out line);
            }
            catch
            {
                file = null;
                line = 0;
                return false;
            }
        }

        private static bool TryGetFileInfoFromFrameInfo(IDebugStackFrame2 debugFrame, out string file, out int line)
        {
            file = null;
            line = 0;
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
                return TryParseFileAndLineFromText(funcName, out file, out line);
            }
            catch
            {
                file = null;
                line = 0;
                return false;
            }
        }

        private static bool TryGetLineNumberFromFrameInfo(IDebugStackFrame2 debugFrame, out int line)
        {
            line = 0;
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

                return TryParseLineFromText(info[0].m_bstrFuncName, out line);
            }
            catch
            {
                line = 0;
                return false;
            }
        }

        private static bool TryParseFileAndLineFromText(string text, out string file, out int line)
        {
            file = null;
            line = 0;
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
                file = match.Groups["file"].Value;
                int.TryParse(match.Groups["line"].Value, out line);
                return !string.IsNullOrEmpty(file);
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
                file = match.Groups["file"].Value;
                return !string.IsNullOrEmpty(file);
            }

            return false;
        }

        private static bool TryParseLineFromText(string text, out int line)
        {
            line = 0;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            Match match = Regex.Match(text, @"\bline\s+(?<line>\d+)\b", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                match = Regex.Match(text, @"\((?<line>\d+)\)\s*$");
            }
            if (!match.Success)
            {
                match = Regex.Match(text, @":(?<line>\d+)\s*$");
            }

            if (!match.Success)
            {
                return false;
            }

            return int.TryParse(match.Groups["line"].Value, out line) && line > 0;
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

        private static bool TryGetFileInfoFromDocumentContext(IDebugDocumentContext2 context, out string file, out int line)
        {
            file = null;
            line = 0;
            if (context == null)
            {
                return false;
            }

            string fullName = null;
            if (Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetName(enum_GETNAME_TYPE.GN_FILENAME, out fullName)))
            {
                file = fullName;
            }

            if (string.IsNullOrEmpty(file) && Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetName(enum_GETNAME_TYPE.GN_URL, out string urlName)))
            {
                file = urlName;
            }

            if (string.IsNullOrEmpty(file) && Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetName(enum_GETNAME_TYPE.GN_BASENAME, out string baseName)))
            {
                file = baseName;
            }
            if (string.IsNullOrEmpty(file) && Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetName(enum_GETNAME_TYPE.GN_NAME, out string name)))
            {
                file = name;
            }
            if (string.IsNullOrEmpty(file) && Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetName(enum_GETNAME_TYPE.GN_MONIKERNAME, out string moniker)))
            {
                file = moniker;
            }
            if (string.IsNullOrEmpty(file) && Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetName(enum_GETNAME_TYPE.GN_TITLE, out string title)))
            {
                file = title;
            }

            var begin = new TEXT_POSITION[1];
            var end = new TEXT_POSITION[1];
            if (Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetStatementRange(begin, end)))
            {
                uint lineValue = begin[0].dwLine;
                if (lineValue != uint.MaxValue)
                {
                    line = unchecked((int)lineValue) + 1;
                }
            }
            if (line <= 0 && Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetSourceRange(begin, end)))
            {
                uint lineValue = begin[0].dwLine;
                if (lineValue != uint.MaxValue)
                {
                    line = unchecked((int)lineValue) + 1;
                }
            }

            return !string.IsNullOrEmpty(file);
        }

        private static bool TryGetLineNumberFromDocumentContext(IDebugDocumentContext2 context, out int line)
        {
            line = 0;
            if (context == null)
            {
                return false;
            }

            var begin = new TEXT_POSITION[1];
            var end = new TEXT_POSITION[1];
            if (Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetStatementRange(begin, end)))
            {
                uint lineValue = begin[0].dwLine;
                if (lineValue != uint.MaxValue)
                {
                    line = unchecked((int)lineValue) + 1;
                    return line > 0;
                }
            }

            if (Microsoft.VisualStudio.ErrorHandler.Succeeded(context.GetSourceRange(begin, end)))
            {
                uint lineValue = begin[0].dwLine;
                if (lineValue != uint.MaxValue)
                {
                    line = unchecked((int)lineValue) + 1;
                    return line > 0;
                }
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

        private void InitializePalette(out Brush namespaceBrush, out Brush functionBrush, out Brush paramNameBrush, out Brush paramValueBrush, out Brush fileBrush, out Brush lineBrush, out Brush punctuationBrush)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var background = Application.Current.TryFindResource(VsBrushes.WindowKey) as SolidColorBrush ?? SystemColors.WindowBrush;
            bool isDark = IsDarkColor(background.Color);
            if (_themeMode == CallStackThemeMode.Light)
            {
                isDark = false;
            }
            else if (_themeMode == CallStackThemeMode.Dark)
            {
                isDark = true;
            }

            Palette palette = isDark ? _darkPalette : _lightPalette;
            namespaceBrush = CreateBrush(palette.Namespace);
            functionBrush = CreateBrush(palette.Function);
            paramNameBrush = CreateBrush(palette.ParamName);
            paramValueBrush = CreateBrush(palette.ParamValue);
            fileBrush = CreateBrush(palette.File);
            lineBrush = CreateBrush(palette.Line);
            punctuationBrush = CreateBrush(palette.Punctuation);
        }

        private void UpdateThemeResources()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _currentFrameBrush = (Brush)Application.Current.TryFindResource(SystemColors.ControlLightBrushKey) ?? Brushes.Transparent;
            InitializePalette(out _namespaceBrush, out _functionBrush, out _paramNameBrush, out _paramValueBrush, out _fileBrush, out _lineBrush, out _punctuationBrush);
        }

        private static bool IsDarkColor(Color color)
        {
            double luminance = (0.299 * color.R) + (0.587 * color.G) + (0.114 * color.B);
            return luminance < 128;
        }

        private static Brush CreateBrush(Color color)
        {
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }

        internal static readonly Palette DarkPalette = new Palette(
            namespaceColor: Color.FromRgb(156, 163, 175),
            functionColor: Color.FromRgb(86, 156, 214),
            paramNameColor: Color.FromRgb(206, 145, 120),
            paramValueColor: Color.FromRgb(181, 206, 168),
            fileColor: Color.FromRgb(156, 220, 254),
            lineColor: Color.FromRgb(255, 198, 109),
            punctuationColor: Color.FromRgb(208, 208, 208));

        internal static readonly Palette LightPalette = new Palette(
            namespaceColor: Color.FromRgb(90, 96, 104),
            functionColor: Color.FromRgb(0, 82, 153),
            paramNameColor: Color.FromRgb(122, 63, 0),
            paramValueColor: Color.FromRgb(0, 100, 0),
            fileColor: Color.FromRgb(28, 98, 139),
            lineColor: Color.FromRgb(170, 75, 0),
            punctuationColor: Color.FromRgb(33, 37, 41));

        private void AddRun(TextBlock textBlock, string text, Brush brush, double fontScale = 1.0, FontFamily fontFamily = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var run = new Run(text);
            if (brush != null)
            {
                run.Foreground = brush;
            }

            if (fontScale > 0 && Math.Abs(fontScale - 1.0) > 0.001)
            {
                run.FontSize = Math.Max(1.0, _fontSize * fontScale);
            }

            if (fontFamily != null)
            {
                run.FontFamily = fontFamily;
            }

            textBlock.Inlines.Add(run);
        }

        private static double ClampScale(double scale)
        {
            if (scale < 0.5)
            {
                return 0.5;
            }

            if (scale > 3.0)
            {
                return 3.0;
            }

            return scale;
        }

        private static string NormalizeFontFamilyName(string fontFamily)
        {
            return string.IsNullOrWhiteSpace(fontFamily) ? null : fontFamily.Trim();
        }

        private static FontFamily CreateFontFamilyOrNull(string fontFamily)
        {
            if (string.IsNullOrWhiteSpace(fontFamily))
            {
                return null;
            }

            try
            {
                return new FontFamily(fontFamily);
            }
            catch
            {
                return null;
            }
        }

        private static void SplitFunctionName(string functionName, out string namespacePart, out string functionPart)
        {
            namespacePart = null;
            functionPart = functionName ?? "<unknown>";
            if (string.IsNullOrEmpty(functionName))
            {
                return;
            }

            int lastDoubleColon = functionName.LastIndexOf("::", StringComparison.Ordinal);
            if (lastDoubleColon >= 0)
            {
                namespacePart = functionName.Substring(0, lastDoubleColon + 2);
                functionPart = functionName.Substring(lastDoubleColon + 2);
                return;
            }

            int lastDot = functionName.LastIndexOf('.');
            if (lastDot >= 0)
            {
                namespacePart = functionName.Substring(0, lastDot + 1);
                functionPart = functionName.Substring(lastDot + 1);
                return;
            }

            int bang = functionName.LastIndexOf('!');
            if (bang >= 0)
            {
                namespacePart = functionName.Substring(0, bang + 1);
                functionPart = functionName.Substring(bang + 1);
            }
        }

        private void FramesList_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = FindAncestor<ListBoxItem>(e.OriginalSource as DependencyObject);
            if (item?.Tag is FrameRowInfo info && info.Frame != null)
            {
                FrameActivated?.Invoke(this, info.Frame);
                e.Handled = true;
            }
        }

        private void FramesList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateSelectionBackgrounds(e.RemovedItems);
            UpdateSelectionBackgrounds(e.AddedItems);
        }

        private void FramesList_OnContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            UpdateContextMenuChecks();
        }

        private void MenuShowNamespace_OnClick(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            bool isChecked = MenuShowNamespace?.IsChecked ?? false;
            if (_showNamespace == isChecked)
            {
                return;
            }

            _showNamespace = isChecked;
            UpdateCallStackIfAvailable(immediateDetails: true);
            DisplayOptionsChanged?.Invoke(this, DisplayOptionsChangedEventArgs.ForShowNamespace(isChecked));
        }

        private void MenuShowParameterTypes_OnClick(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            bool isChecked = MenuShowParameterTypes?.IsChecked ?? false;
            if (_showParameterTypes == isChecked)
            {
                return;
            }

            _showParameterTypes = isChecked;
            UpdateCallStackIfAvailable(immediateDetails: true);
            DisplayOptionsChanged?.Invoke(this, DisplayOptionsChangedEventArgs.ForShowParameterTypes(isChecked));
        }

        private void MenuShowLineNumbers_OnClick(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            bool isChecked = MenuShowLineNumbers?.IsChecked ?? false;
            if (_showLineNumbers == isChecked)
            {
                return;
            }

            _showLineNumbers = isChecked;
            UpdateCallStackIfAvailable(immediateDetails: true);
            DisplayOptionsChanged?.Invoke(this, DisplayOptionsChangedEventArgs.ForShowLineNumbers(isChecked));
        }

        private void MenuShowFilePath_OnClick(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            bool isChecked = MenuShowFilePath?.IsChecked ?? false;
            if (_showFilePath == isChecked)
            {
                return;
            }

            _showFilePath = isChecked;
            UpdateCallStackIfAvailable(immediateDetails: true);
            DisplayOptionsChanged?.Invoke(this, DisplayOptionsChangedEventArgs.ForShowFilePath(isChecked));
        }

        private void MenuResetFactoryDefaults_OnClick(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _skipNextDetailsDelay = true;
            ResetRequested?.Invoke(this, EventArgs.Empty);
        }

        private void MenuFontIncrease_OnClick(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            RequestFontSizeStep(1);
        }

        private void MenuFontDecrease_OnClick(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            RequestFontSizeStep(-1);
        }

        private void FramesList_OnPreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!IsKeyboardFocusWithin)
            {
                return;
            }

            if ((Keyboard.Modifiers & ModifierKeys.Control) == 0)
            {
                return;
            }

            if (e.Delta > 0)
            {
                RequestFontSizeStep(1);
                e.Handled = true;
                return;
            }

            if (e.Delta < 0)
            {
                RequestFontSizeStep(-1);
                e.Handled = true;
            }
        }

        private void FramesList_OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!IsKeyboardFocusWithin)
            {
                return;
            }

            if ((Keyboard.Modifiers & ModifierKeys.Control) == 0)
            {
                return;
            }

            if (e.Key == Key.OemPlus || e.Key == Key.Add || e.Key == Key.OemComma)
            {
                RequestFontSizeStep(1);
                e.Handled = true;
                return;
            }

            if (e.Key == Key.OemMinus || e.Key == Key.Subtract)
            {
                RequestFontSizeStep(-1);
                e.Handled = true;
            }
        }

        private void MenuDisplayHex_OnClick(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            SetHexDisplayMode(true);
        }

        private void MenuDisplayDecimal_OnClick(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            SetHexDisplayMode(false);
        }

        private void SetHexDisplayMode(bool hexDisplayMode)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_hexDisplayMode == hexDisplayMode)
            {
                UpdateContextMenuChecks();
                return;
            }

            _hexDisplayMode = hexDisplayMode;
            UpdateContextMenuChecks();
            DisplayOptionsChanged?.Invoke(this, DisplayOptionsChangedEventArgs.ForHexDisplayMode(hexDisplayMode));
        }

        private void RequestFontSizeStep(int step)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (step == 0)
            {
                return;
            }

            _skipNextDetailsDelay = true;
            FontSizeStepRequested?.Invoke(this, step);
        }

        private void UpdateCallStackIfAvailable(bool immediateDetails = false)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_lastFrames != null)
            {
                if (immediateDetails)
                {
                    _skipNextDetailsDelay = true;
                }

                UpdateCallStack(_lastFrames, _lastCurrentFrame);
            }
        }

        private void UpdateContextMenuChecks()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (MenuShowNamespace != null)
            {
                MenuShowNamespace.IsChecked = _showNamespace;
            }

            if (MenuShowParameterTypes != null)
            {
                MenuShowParameterTypes.IsChecked = _showParameterTypes;
            }

            if (MenuShowLineNumbers != null)
            {
                MenuShowLineNumbers.IsChecked = _showLineNumbers;
            }

            if (MenuShowFilePath != null)
            {
                MenuShowFilePath.IsChecked = _showFilePath;
            }

            if (MenuDisplayHex != null)
            {
                MenuDisplayHex.IsChecked = _hexDisplayMode;
            }

            if (MenuDisplayDecimal != null)
            {
                MenuDisplayDecimal.IsChecked = !_hexDisplayMode;
            }
        }

        private void UpdateSelectionBackgrounds(System.Collections.IList items)
        {
            foreach (var item in items)
            {
                if (item is ListBoxItem listItem && listItem.Tag is FrameRowInfo info)
                {
                    info.UpdateBackground(listItem.IsSelected, _currentFrameBrush);
                }
            }
        }

        private void BeginPopulateDetails()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            CancelDetailsPopulation();

            var items = new List<ListBoxItem>(FramesList.Items.Count);
            foreach (var item in FramesList.Items)
            {
                if (item is ListBoxItem listItem)
                {
                    items.Add(listItem);
                }
            }

            if (items.Count == 0)
            {
                return;
            }

            _detailsCts = new CancellationTokenSource();
            CancellationToken token = _detailsCts.Token;
            int delayMs = _skipNextDetailsDelay ? 0 : 200;
            _skipNextDetailsDelay = false;
            PopulateDetailsAsync(items, token, delayMs).FileAndForget("ColorCallStack/PopulateDetails");
        }

        private void CancelDetailsPopulation()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_detailsCts != null)
            {
                _detailsCts.Cancel();
                _detailsCts.Dispose();
                _detailsCts = null;
            }
        }

        private async Task PopulateDetailsAsync(IReadOnlyList<ListBoxItem> items, CancellationToken token, int delayMs)
        {
            if (delayMs > 0)
            {
                await Task.Delay(delayMs, token);
            }
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            foreach (var item in items)
            {
                token.ThrowIfCancellationRequested();
                if (item?.Tag is FrameRowInfo info)
                {
                    try
                    {
                        if (info.FunctionText != null)
                        {
                            info.FunctionText.Inlines.Clear();
                            BuildFrameInlines(info.FunctionText, info.Frame, includeArguments: true);
                        }

                        if (info.FileText != null)
                        {
                            info.FileText.Inlines.Clear();
                            BuildFileInlines(info.FileText, info.Frame, includeFileInfo: true);
                        }
                    }
                    catch
                    {
                        // Ignore per-frame failures to keep UI responsive.
                    }
                }

                await Task.Yield();
            }
        }

        private static T FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T typed)
                {
                    return typed;
                }

                if (current is Visual || current is System.Windows.Media.Media3D.Visual3D)
                {
                    current = VisualTreeHelper.GetParent(current);
                }
                else if (current is FrameworkContentElement contentElement)
                {
                    current = contentElement.Parent ?? LogicalTreeHelper.GetParent(contentElement);
                }
                else
                {
                    current = LogicalTreeHelper.GetParent(current);
                }
            }

            return null;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_themeHooked)
            {
                return;
            }

            VSColorTheme.ThemeChanged += VSColorTheme_ThemeChanged;
            _themeHooked = true;
        }

        private void OnUnloaded(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!_themeHooked)
            {
                return;
            }

            VSColorTheme.ThemeChanged -= VSColorTheme_ThemeChanged;
            _themeHooked = false;
        }

        private void VSColorTheme_ThemeChanged(ThemeChangedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            UpdateThemeResources();
            if (_lastFrames != null)
            {
                UpdateCallStack(_lastFrames, _lastCurrentFrame);
            }
        }

        private sealed class FrameRowInfo
        {
            public FrameRowInfo(StackFrame frame, Border border, TextBlock functionText, TextBlock fileText, bool isCurrent)
            {
                Frame = frame;
                Border = border;
                FunctionText = functionText;
                FileText = fileText;
                IsCurrent = isCurrent;
            }

            public StackFrame Frame { get; }
            public Border Border { get; }
            public TextBlock FunctionText { get; }
            public TextBlock FileText { get; }
            public bool IsCurrent { get; }

            public void UpdateBackground(bool isSelected, Brush currentBrush)
            {
                if (!IsCurrent)
                {
                    Border.Background = Brushes.Transparent;
                    return;
                }

                Border.Background = isSelected ? Brushes.Transparent : currentBrush;
            }
        }

        internal readonly struct Palette
        {
            public Palette(Color namespaceColor, Color functionColor, Color paramNameColor, Color paramValueColor, Color fileColor, Color lineColor, Color punctuationColor)
            {
                Namespace = namespaceColor;
                Function = functionColor;
                ParamName = paramNameColor;
                ParamValue = paramValueColor;
                File = fileColor;
                Line = lineColor;
                Punctuation = punctuationColor;
            }

            public Color Namespace { get; }
            public Color Function { get; }
            public Color ParamName { get; }
            public Color ParamValue { get; }
            public Color File { get; }
            public Color Line { get; }
            public Color Punctuation { get; }
        }

        private struct ArgumentPart
        {
            public ArgumentPart(string name, string value, string type)
            {
                Name = name;
                Value = value;
                Type = type;
            }

            public string Name { get; }
            public string Value { get; }
            public string Type { get; }
        }

        internal sealed class DisplayOptionsChangedEventArgs : EventArgs
        {
            private DisplayOptionsChangedEventArgs(bool? showNamespace, bool? showParameterTypes, bool? showLineNumbers, bool? showFilePath, bool? hexDisplayMode)
            {
                ShowNamespace = showNamespace;
                ShowParameterTypes = showParameterTypes;
                ShowLineNumbers = showLineNumbers;
                ShowFilePath = showFilePath;
                HexDisplayMode = hexDisplayMode;
            }

            public bool? ShowNamespace { get; }
            public bool? ShowParameterTypes { get; }
            public bool? ShowLineNumbers { get; }
            public bool? ShowFilePath { get; }
            public bool? HexDisplayMode { get; }

            public static DisplayOptionsChangedEventArgs ForShowNamespace(bool value) => new DisplayOptionsChangedEventArgs(value, null, null, null, null);
            public static DisplayOptionsChangedEventArgs ForShowParameterTypes(bool value) => new DisplayOptionsChangedEventArgs(null, value, null, null, null);
            public static DisplayOptionsChangedEventArgs ForShowLineNumbers(bool value) => new DisplayOptionsChangedEventArgs(null, null, value, null, null);
            public static DisplayOptionsChangedEventArgs ForShowFilePath(bool value) => new DisplayOptionsChangedEventArgs(null, null, null, value, null);
            public static DisplayOptionsChangedEventArgs ForHexDisplayMode(bool value) => new DisplayOptionsChangedEventArgs(null, null, null, null, value);
        }
    }
}
