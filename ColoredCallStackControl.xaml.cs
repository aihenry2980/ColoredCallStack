using EnvDTE;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

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
        private bool _themeHooked;
        private List<StackFrame> _lastFrames;
        private StackFrame _lastCurrentFrame;

        internal event EventHandler<StackFrame> FrameActivated;

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

        public void ClearCallStack()
        {
            FramesList.Items.Clear();
        }

        public void UpdateCallStack(IReadOnlyList<StackFrame> frames, StackFrame currentFrame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _lastFrames = frames != null ? new List<StackFrame>(frames) : null;
            _lastCurrentFrame = currentFrame;
            FramesList.Items.Clear();
            if (frames == null || frames.Count == 0)
            {
                return;
            }

            ListBoxItem currentItem = null;
            foreach (var frame in frames)
            {
                bool isCurrent = IsSameFrame(frame, currentFrame);
                var item = CreateFrameItem(frame, isCurrent);
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
                FontFamily = new FontFamily("Consolas"),
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping = TextWrapping.NoWrap
            };
            BuildFrameInlines(text, frame);
            Grid.SetColumn(text, 1);

            grid.Children.Add(arrow);
            grid.Children.Add(text);
            border.Child = grid;

            var item = new ListBoxItem
            {
                Content = border,
                Tag = new FrameRowInfo(frame, border, isCurrent)
            };

            return item;
        }

        private void BuildFrameInlines(TextBlock textBlock, StackFrame frame)
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

            string function = frame.FunctionName ?? "<unknown>";
            SplitFunctionName(function, out string namespacePart, out string functionPart);

            if (!string.IsNullOrEmpty(namespacePart))
            {
                AddRun(textBlock, namespacePart, _namespaceBrush);
            }

            AddRun(textBlock, functionPart, _functionBrush);

            if (TryGetArguments(frame.Arguments, out List<ArgumentPart> args) && args.Count > 0)
            {
                AddRun(textBlock, "(", _punctuationBrush);
                for (int i = 0; i < args.Count; i++)
                {
                    if (i > 0)
                    {
                        AddRun(textBlock, ", ", _punctuationBrush);
                    }

                    ArgumentPart arg = args[i];
                    if (!string.IsNullOrEmpty(arg.Name))
                    {
                        AddRun(textBlock, arg.Name, _paramNameBrush);
                        if (!string.IsNullOrEmpty(arg.Value))
                        {
                            AddRun(textBlock, "=", _punctuationBrush);
                            AddRun(textBlock, arg.Value, _paramValueBrush);
                        }
                    }
                    else if (!string.IsNullOrEmpty(arg.Value))
                    {
                        AddRun(textBlock, arg.Value, _paramValueBrush);
                    }
                }
                AddRun(textBlock, ")", _punctuationBrush);
            }

            if (TryGetFileInfo(frame, out string file, out int line))
            {
                AddRun(textBlock, "  ", _punctuationBrush);
                AddRun(textBlock, System.IO.Path.GetFileName(file), _fileBrush);
                AddRun(textBlock, ":", _punctuationBrush);
                AddRun(textBlock, line.ToString(), _lineBrush);
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
                    parts.Add(new ArgumentPart(name, value));
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

        private static bool TryGetFileInfo(StackFrame frame, out string file, out int line)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            file = null;
            line = 0;
            if (frame == null)
            {
                return false;
            }

            try
            {
                dynamic dyn = frame;
                file = dyn.FileName as string;
                line = (int)dyn.LineNumber;
                return !string.IsNullOrEmpty(file) && line > 0;
            }
            catch
            {
                file = null;
                line = 0;
                return false;
            }
        }

        private void InitializePalette(out Brush namespaceBrush, out Brush functionBrush, out Brush paramNameBrush, out Brush paramValueBrush, out Brush fileBrush, out Brush lineBrush, out Brush punctuationBrush)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var background = Application.Current.TryFindResource(VsBrushes.WindowKey) as SolidColorBrush ?? SystemColors.WindowBrush;
            bool isDark = IsDarkColor(background.Color);

            Palette palette = isDark ? DarkPalette : LightPalette;
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

        private static readonly Palette DarkPalette = new Palette(
            namespaceColor: Color.FromRgb(156, 163, 175),
            functionColor: Color.FromRgb(86, 156, 214),
            paramNameColor: Color.FromRgb(206, 145, 120),
            paramValueColor: Color.FromRgb(181, 206, 168),
            fileColor: Color.FromRgb(156, 220, 254),
            lineColor: Color.FromRgb(255, 198, 109),
            punctuationColor: Color.FromRgb(208, 208, 208));

        private static readonly Palette LightPalette = new Palette(
            namespaceColor: Color.FromRgb(90, 96, 104),
            functionColor: Color.FromRgb(0, 82, 153),
            paramNameColor: Color.FromRgb(122, 63, 0),
            paramValueColor: Color.FromRgb(0, 100, 0),
            fileColor: Color.FromRgb(28, 98, 139),
            lineColor: Color.FromRgb(170, 75, 0),
            punctuationColor: Color.FromRgb(33, 37, 41));

        private void AddRun(TextBlock textBlock, string text, Brush brush)
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
            textBlock.Inlines.Add(run);
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

        private static T FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T typed)
                {
                    return typed;
                }

                current = VisualTreeHelper.GetParent(current);
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
            public FrameRowInfo(StackFrame frame, Border border, bool isCurrent)
            {
                Frame = frame;
                Border = border;
                IsCurrent = isCurrent;
            }

            public StackFrame Frame { get; }
            public Border Border { get; }
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

        private readonly struct Palette
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
            public ArgumentPart(string name, string value)
            {
                Name = name;
                Value = value;
            }

            public string Name { get; }
            public string Value { get; }
        }
    }
}
