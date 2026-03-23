using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace LeadMailer.Behaviors;

public static class RichTextBoxBindingBehavior
{
    public static readonly DependencyProperty DocumentXamlProperty =
        DependencyProperty.RegisterAttached(
            "DocumentXaml",
            typeof(string),
            typeof(RichTextBoxBindingBehavior),
            new FrameworkPropertyMetadata(null,
                FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
                OnDocumentXamlChanged));

    private static readonly DependencyProperty IsInternalUpdateProperty =
        DependencyProperty.RegisterAttached("IsInternalUpdate", typeof(bool), typeof(RichTextBoxBindingBehavior));

    public static string GetDocumentXaml(DependencyObject obj) => (string?)obj.GetValue(DocumentXamlProperty) ?? string.Empty;
    public static void SetDocumentXaml(DependencyObject obj, string value) => obj.SetValue(DocumentXamlProperty, value);

    private static bool GetIsInternalUpdate(DependencyObject obj) => (bool)obj.GetValue(IsInternalUpdateProperty);
    private static void SetIsInternalUpdate(DependencyObject obj, bool value) => obj.SetValue(IsInternalUpdateProperty, value);

    private static void OnDocumentXamlChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not RichTextBox rtb) return;

        rtb.TextChanged -= RichTextBoxOnTextChanged;

        if (!GetIsInternalUpdate(rtb))
        {
            SetIsInternalUpdate(rtb, true);
            rtb.Document = ToDocument(e.NewValue as string);
            SetIsInternalUpdate(rtb, false);
        }

        rtb.TextChanged += RichTextBoxOnTextChanged;
    }

    private static void RichTextBoxOnTextChanged(object sender, TextChangedEventArgs e)
    {
        if (sender is not RichTextBox rtb || GetIsInternalUpdate(rtb)) return;

        SetIsInternalUpdate(rtb, true);
        SetDocumentXaml(rtb, FromDocument(rtb.Document));
        SetIsInternalUpdate(rtb, false);
    }

    private static FlowDocument ToDocument(string? xaml)
    {
        var doc = new FlowDocument();
        if (string.IsNullOrWhiteSpace(xaml)) return doc;

        var range = new TextRange(doc.ContentStart, doc.ContentEnd);
        try
        {
            using var ms = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xaml));
            range.Load(ms, DataFormats.Xaml);
        }
        catch
        {
            range.Text = xaml;
        }

        return doc;
    }

    private static string FromDocument(FlowDocument doc)
    {
        var range = new TextRange(doc.ContentStart, doc.ContentEnd);
        using var ms = new MemoryStream();
        range.Save(ms, DataFormats.Xaml);
        return System.Text.Encoding.UTF8.GetString(ms.ToArray());
    }
}
