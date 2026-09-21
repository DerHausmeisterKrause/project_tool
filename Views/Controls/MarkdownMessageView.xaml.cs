using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace TaskTool.Views.Controls;

public partial class MarkdownMessageView : UserControl
{
    public static readonly DependencyProperty MarkdownProperty = DependencyProperty.Register(nameof(Markdown), typeof(string), typeof(MarkdownMessageView), new PropertyMetadata(string.Empty, (dependencyObject, _) => ((MarkdownMessageView)dependencyObject).Render()));
    private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
    public string Markdown { get => (string)GetValue(MarkdownProperty); set => SetValue(MarkdownProperty, value); }
    public MarkdownMessageView() { InitializeComponent(); Render(); }

    private void Render()
    {
        var document = new FlowDocument { PagePadding = new Thickness(0), Background = Brushes.Transparent, FontFamily = FontFamily, FontSize = FontSize, Foreground = Foreground, TextAlignment = TextAlignment.Left };
        if (!string.IsNullOrEmpty(Markdown)) foreach (var block in Markdig.Markdown.Parse(Markdown, Pipeline)) AddBlock(document, block);
        Viewer.Document = document;
    }
    private static void AddBlock(FlowDocument document, Block block)
    {
        if (block is CodeBlock codeBlock)
        {
            var code = new TextBox { Text = codeBlock.Lines.ToString(), IsReadOnly = true, AcceptsReturn = true, TextWrapping = TextWrapping.Wrap, FontFamily = new FontFamily("Consolas"), Background = new SolidColorBrush(Color.FromRgb(15, 23, 42)), Foreground = new SolidColorBrush(Color.FromRgb(226, 232, 240)), BorderThickness = new Thickness(0), Padding = new Thickness(12) };
            document.Blocks.Add(new BlockUIContainer(new Border { Child = code, Background = code.Background, CornerRadius = new CornerRadius(6), Margin = new Thickness(0, 5, 0, 9) })); return;
        }
        if (block is Markdig.Syntax.ListBlock list)
        {
            var wpfList = new System.Windows.Documents.List { MarkerStyle = list.IsOrdered ? TextMarkerStyle.Decimal : TextMarkerStyle.Disc, Margin = new Thickness(18, 4, 0, 8) };
            foreach (ListItemBlock item in list) { var li = new ListItem(); foreach (var child in item) li.Blocks.Add(CreateParagraph(child)); wpfList.ListItems.Add(li); }
            document.Blocks.Add(wpfList); return;
        }
        var paragraph = CreateParagraph(block);
        if (block is HeadingBlock heading) { paragraph.FontSize = heading.Level switch { 1 => 24, 2 => 21, 3 => 18, _ => 16 }; paragraph.FontWeight = FontWeights.SemiBold; paragraph.Margin = new Thickness(0, 9, 0, 6); }
        document.Blocks.Add(paragraph);
    }
    private static Paragraph CreateParagraph(Block block)
    {
        var paragraph = new Paragraph { Margin = new Thickness(0, 0, 0, 9) };
        if (block is LeafBlock { Inline: { } inline }) AddInlines(paragraph.Inlines, inline.FirstChild);
        else paragraph.Inlines.Add(new Run(block.ToString()));
        return paragraph;
    }
    private static void AddInlines(InlineCollection target, Markdig.Syntax.Inlines.Inline? current)
    {
        while (current != null)
        {
            switch (current)
            {
                case LiteralInline literal: target.Add(new Run(literal.Content.ToString())); break;
                case CodeInline code: target.Add(new Run(code.Content) { FontFamily = new FontFamily("Consolas"), Background = new SolidColorBrush(Color.FromRgb(30, 41, 59)) }); break;
                case LineBreakInline: target.Add(new LineBreak()); break;
                case EmphasisInline emphasis:
                    var span = new Span(); AddInlines(span.Inlines, emphasis.FirstChild); if (emphasis.DelimiterCount >= 2) span.FontWeight = FontWeights.Bold; else span.FontStyle = FontStyles.Italic; target.Add(span); break;
                case ContainerInline container: var nested = new Span(); AddInlines(nested.Inlines, container.FirstChild); target.Add(nested); break;
            }
            current = current.NextSibling;
        }
    }
}
