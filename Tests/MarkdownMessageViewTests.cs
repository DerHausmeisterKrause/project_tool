using System.Windows.Documents;
using TaskTool.Views.Controls;
using Xunit;

namespace TaskTool.Tests;

public sealed class MarkdownMessageViewTests
{
    [Fact]
    public void NestedOrderedAndUnorderedLists_AreRenderedRecursively()
    {
        const string markdown = "1. Systemeinstellungen\n   - Autostart prüfen\n   - Updates prüfen\n\n2. Hardware prüfen\n   1. CPU\n   2. RAM\n   3. Datenträger";
        var document = MarkdownMessageView.RenderDocument(markdown);
        var outer = Assert.IsType<List>(Assert.Single(document.Blocks));
        Assert.Equal(2, outer.ListItems.Count);
        Assert.Contains(outer.ListItems.Cast<ListItem>(), item => item.Blocks.OfType<List>().Any(nested => nested.MarkerStyle == System.Windows.TextMarkerStyle.Disc));
        Assert.Contains(outer.ListItems.Cast<ListItem>(), item => item.Blocks.OfType<List>().Any(nested => nested.MarkerStyle == System.Windows.TextMarkerStyle.Decimal));
        var text = new TextRange(document.ContentStart, document.ContentEnd).Text;
        Assert.DoesNotContain("Markdig.Syntax", text);
        Assert.Contains("Autostart prüfen", text); Assert.Contains("Datenträger", text);
    }
}
