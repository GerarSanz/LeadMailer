using LeadMailer.Services;
using OfficeOpenXml;
using System.Windows;
using System.Windows.Threading;

namespace LeadMailer;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        ConfigureEpplusLicense();
        RegisterExceptionHandlers();
        AppLogger.Info($"Aplicación iniciada. Versión {GetVersion()}");
        base.OnStartup(e);
    }

    protected override void OnExit(ExitEventArgs e)
    {
        AppLogger.Info("Aplicación cerrada.");
        base.OnExit(e);
    }

    // ── Excepciones no controladas ────────────────────────────────────────────

    private void RegisterExceptionHandlers()
    {
        // Excepciones en el hilo UI (binding errors, command handlers, etc.)
        DispatcherUnhandledException += OnDispatcherUnhandledException;

        // Excepciones en tareas async no observadas
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;

        // Excepciones fatales en hilos de fondo
        AppDomain.CurrentDomain.UnhandledException += OnAppDomainUnhandledException;
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        AppLogger.Error("Excepción no controlada (UI)", e.Exception);
        MessageBox.Show(
            $"Se produjo un error inesperado:\n\n{e.Exception.Message}\n\n" +
            $"El error ha sido registrado en:\n%AppData%\\LeadMailer\\logs\\",
            "LeadMailer — Error",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
        e.Handled = true;
    }

    private static void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        AppLogger.Error("Excepción no observada en Task", e.Exception);
        e.SetObserved();
    }

    private static void OnAppDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        var ex = e.ExceptionObject as Exception;
        AppLogger.Error("Excepción fatal (AppDomain)", ex!);
        if (e.IsTerminating)
            MessageBox.Show(
                $"Error crítico. La aplicación debe cerrarse.\n\n{ex?.Message}\n\n" +
                $"Consulta el log en %AppData%\\LeadMailer\\logs\\",
                "LeadMailer — Error Crítico",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static void ConfigureEpplusLicense()
    {
        ExcelPackage.License.SetNonCommercialPersonal("LeadMailer");
    }

    private static string GetVersion()
        => System.Reflection.Assembly.GetExecutingAssembly()
               .GetName().Version?.ToString() ?? "desconocida";
}

