using LeadMailer.ViewModels;
using System.Windows;
using System.Windows.Controls;

namespace LeadMailer.Views;

public partial class MainWindow : Window
{
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
    }
}
