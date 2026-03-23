using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using LeadMailer.Models;

namespace LeadMailer.Converters
{
    [ValueConversion(typeof(bool), typeof(Visibility))]
    public class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type t, object p, CultureInfo c)
            => value is true ? Visibility.Visible : Visibility.Collapsed;
        public object ConvertBack(object value, Type t, object p, CultureInfo c)
            => value is Visibility.Visible;
    }

    [ValueConversion(typeof(bool), typeof(Visibility))]
    public class InvertBoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type t, object p, CultureInfo c)
            => value is true ? Visibility.Collapsed : Visibility.Visible;
        public object ConvertBack(object value, Type t, object p, CultureInfo c)
            => throw new NotImplementedException();
    }

    /// <summary>true → verde (enviado) | false → gris (pendiente)</summary>
    [ValueConversion(typeof(bool), typeof(Brush))]
    public class SentStatusBrushConverter : IValueConverter
    {
        public object Convert(object value, Type t, object p, CultureInfo c)
            => value is true
                ? new SolidColorBrush(Color.FromRgb(22, 163, 74))   // green-600
                : new SolidColorBrush(Color.FromRgb(100, 116, 139)); // slate-500
        public object ConvertBack(object value, Type t, object p, CultureInfo c)
            => throw new NotImplementedException();
    }

    /// <summary>null → Collapsed | not null → Visible</summary>
    public class NullToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type t, object p, CultureInfo c)
            => value != null ? Visibility.Visible : Visibility.Collapsed;
        public object ConvertBack(object value, Type t, object p, CultureInfo c)
            => throw new NotImplementedException();
    }

    /// <summary>string vacío → Collapsed | con texto → Visible</summary>
    public class EmptyStringToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type t, object p, CultureInfo c)
            => !string.IsNullOrWhiteSpace(value as string) ? Visibility.Visible : Visibility.Collapsed;
        public object ConvertBack(object value, Type t, object p, CultureInfo c)
            => throw new NotImplementedException();
    }

    /// <summary>
    /// Compara el valor con el parámetro (string).
    /// true si son iguales → para bindear RadioButton a una propiedad string.
    /// </summary>
    public class EqualStringConverter : IValueConverter
    {
        public object Convert(object value, Type t, object p, CultureInfo c)
            => string.Equals(value?.ToString(), p?.ToString(), StringComparison.OrdinalIgnoreCase);

        public object ConvertBack(object value, Type t, object p, CultureInfo c)
            => value is true ? p?.ToString() ?? "" : Binding.DoNothing;
    }

    /// <summary>LeadStatus → color del badge de estado.</summary>
    [ValueConversion(typeof(LeadStatus), typeof(Brush))]
    public class LeadStatusBrushConverter : IValueConverter
    {
        public object Convert(object value, Type t, object p, CultureInfo c)
            => value is LeadStatus s
                ? s switch
                {
                    LeadStatus.Enviado    => new SolidColorBrush(Color.FromRgb(22,  163,  74)),  // green-600
                    LeadStatus.Descartado => new SolidColorBrush(Color.FromRgb(100, 116, 139)),  // slate-500
                    _                     => new SolidColorBrush(Color.FromRgb(245, 158,  11))   // amber-500
                }
                : new SolidColorBrush(Color.FromRgb(245, 158, 11));
        public object ConvertBack(object value, Type t, object p, CultureInfo c)
            => throw new NotImplementedException();
    }

    /// <summary>IsDiscarded bool → etiqueta del botón Descartar / Recuperar.</summary>
    [ValueConversion(typeof(bool), typeof(string))]
    public class DiscardButtonLabelConverter : IValueConverter
    {
        public object Convert(object value, Type t, object p, CultureInfo c)
            => value is true ? "↩ Recuperar" : "⊘ Descartar";
        public object ConvertBack(object value, Type t, object p, CultureInfo c)
            => throw new NotImplementedException();
    }

    /// <summary>LeadLabel → texto en español.</summary>
    public class LeadLabelTextConverter : IValueConverter
    {
        public object Convert(object value, Type t, object p, CultureInfo c)
            => value is LeadLabel l ? l switch
            {
                LeadLabel.Interesado => "Interesado",
                LeadLabel.Confirmado => "Confirmado",
                LeadLabel.Descartado => "Descartado",
                _                    => "—"
            } : "—";
        public object ConvertBack(object value, Type t, object p, CultureInfo c)
            => throw new NotImplementedException();
    }

    /// <summary>LeadLabel → color de fondo del badge.</summary>
    public class LeadLabelBrushConverter : IValueConverter
    {
        public object Convert(object value, Type t, object p, CultureInfo c)
            => value is LeadLabel l ? l switch
            {
                LeadLabel.Interesado => new SolidColorBrush(Color.FromRgb( 37,  99, 235)),  // blue-600
                LeadLabel.Confirmado => new SolidColorBrush(Color.FromRgb(124,  58, 237)),  // violet-600
                LeadLabel.Descartado => new SolidColorBrush(Color.FromRgb(100, 116, 139)),  // slate-500
                _                    => new SolidColorBrush(Color.FromRgb(226, 232, 240))   // slate-200
            } : new SolidColorBrush(Color.FromRgb(226, 232, 240));
        public object ConvertBack(object value, Type t, object p, CultureInfo c)
            => throw new NotImplementedException();
    }

    /// <summary>LeadLabel → color de texto del badge (blanco para etiquetas con color, gris oscuro para Ninguna).</summary>
    public class LeadLabelForegroundConverter : IValueConverter
    {
        public object Convert(object value, Type t, object p, CultureInfo c)
            => value is LeadLabel.Ninguna
                ? new SolidColorBrush(Color.FromRgb(71, 85, 105))  // slate-600
                : new SolidColorBrush(Colors.White);
        public object ConvertBack(object value, Type t, object p, CultureInfo c)
            => throw new NotImplementedException();
    }
}
