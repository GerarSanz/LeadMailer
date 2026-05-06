using LeadMailer.ViewModels;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace LeadMailer.Views;

public partial class MainWindow : Window
{
    private bool _closingAlreadyHandled;

    public MainWindow()
    {
        InitializeComponent();

        // PasswordBox no soporta binding por seguridad, lo manejamos manualmente
        PwdBox.PasswordChanged += (s, e) =>
        {
            if (DataContext is MainViewModel vm)
                vm.SmtpConfig.Password = PwdBox.Password;
        };

        Loaded += (s, e) =>
        {
            if (DataContext is MainViewModel vm && !string.IsNullOrEmpty(vm.SmtpConfig.Password))
                PwdBox.Password = vm.SmtpConfig.Password;
        };

        Closing += OnMainWindowClosing;
    }

    private void OnMainWindowClosing(object? sender, CancelEventArgs e)
    {
        if (_closingAlreadyHandled)
            return;

        if (DataContext is not MainViewModel vm)
            return;

        var decision = MessageBox.Show(
            "¿Deseas enviar el reporte de leads por email antes de cerrar?",
            "Cerrar aplicación",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Question);

        if (decision == MessageBoxResult.Cancel)
        {
            e.Cancel = true;
            return;
        }

        e.Cancel = true;

        if (decision == MessageBoxResult.No)
        {
            _closingAlreadyHandled = true;
            Dispatcher.BeginInvoke(new Action(Close), DispatcherPriority.Background);
            return;
        }

        _closingAlreadyHandled = true;
        Dispatcher.BeginInvoke(async () =>
        {
            await vm.SendClosingLeadsReportAsync();
            Close();
        }, DispatcherPriority.Background);
    }
}
