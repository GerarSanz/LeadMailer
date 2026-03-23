using System.IO;
using System.Net;
using System.Text;
using System.Windows;
using System.Windows.Documents;

namespace LeadMailer.Helpers;

public static class RichTextSerialization
{
    public static string PlainTextToXaml(string text)
    {
        var doc = new FlowDocument();
        var range = new TextRange(doc.ContentStart, doc.ContentEnd)
        {
            Text = text ?? string.Empty
        };

        using var ms = new MemoryStream();
        range.Save(ms, DataFormats.Xaml);
        return System.Text.Encoding.UTF8.GetString(ms.ToArray());
    }

    public static string XamlToPlainText(string xaml)
    {
        if (string.IsNullOrWhiteSpace(xaml)) return string.Empty;

        var doc = new FlowDocument();
        var range = new TextRange(doc.ContentStart, doc.ContentEnd);

        try
        {
            using var ms = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xaml));
            range.Load(ms, DataFormats.Xaml);
            return range.Text.TrimEnd('\r', '\n');
        }
        catch
        {
            return xaml;
        }
    }

    public static string XamlToHtml(string xaml)
    {
        if (string.IsNullOrWhiteSpace(xaml)) return string.Empty;

        var doc = new FlowDocument();
        var range = new TextRange(doc.ContentStart, doc.ContentEnd);

        try
        {
            using var ms = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xaml));
            range.Load(ms, DataFormats.Xaml);

            var sb = new StringBuilder();
            foreach (var block in doc.Blocks)
                AppendBlockHtml(block, sb);
            return sb.ToString();
        }
        catch
        {
            return $"<p>{WebUtility.HtmlEncode(xaml).Replace("\n", "<br/>")}</p>";
        }
    }

    private static void AppendBlockHtml(Block block, StringBuilder sb)
    {
        switch (block)
        {
            case Paragraph p:
                sb.Append("<p style='margin:0 0 12px'>");
                foreach (var inline in p.Inlines)
                    AppendInlineHtml(inline, sb);
                sb.Append("</p>");
                break;

            case List list:
                sb.Append(list.MarkerStyle == TextMarkerStyle.Decimal ? "<ol style='margin:0 0 12px 20px;padding:0'>" : "<ul style='margin:0 0 12px 20px;padding:0'>");
                foreach (var item in list.ListItems)
                {
                    sb.Append("<li style='margin:0 0 4px'>");
                    foreach (var liBlock in item.Blocks)
                    {
                        if (liBlock is Paragraph lip)
                        {
                            foreach (var inline in lip.Inlines)
                                AppendInlineHtml(inline, sb);
                        }
                        else
                        {
                            AppendBlockHtml(liBlock, sb);
                        }
                    }
                    sb.Append("</li>");
                }
                sb.Append(list.MarkerStyle == TextMarkerStyle.Decimal ? "</ol>" : "</ul>");
                break;
        }
    }

    private static void AppendInlineHtml(Inline inline, StringBuilder sb)
    {
        switch (inline)
        {
            case Run run:
                sb.Append(WebUtility.HtmlEncode(run.Text));
                break;
            case LineBreak:
                sb.Append("<br/>");
                break;
            case Bold bold:
                sb.Append("<strong>");
                foreach (var child in bold.Inlines) AppendInlineHtml(child, sb);
                sb.Append("</strong>");
                break;
            case Italic italic:
                sb.Append("<em>");
                foreach (var child in italic.Inlines) AppendInlineHtml(child, sb);
                sb.Append("</em>");
                break;
            case Underline underline:
                sb.Append("<u>");
                foreach (var child in underline.Inlines) AppendInlineHtml(child, sb);
                sb.Append("</u>");
                break;
            case Hyperlink link:
                var href = link.NavigateUri?.ToString() ?? "#";
                sb.Append($"<a href='{WebUtility.HtmlEncode(href)}'>");
                foreach (var child in link.Inlines) AppendInlineHtml(child, sb);
                sb.Append("</a>");
                break;
            case Span span:
                foreach (var child in span.Inlines) AppendInlineHtml(child, sb);
                break;
        }
    }
}
