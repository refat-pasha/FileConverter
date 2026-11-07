using System.Windows;
using FileConverter.Services;

namespace FileConverter;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Headless CLI mode when any argument is given; otherwise open the GUI.
        if (e.Args.Length > 0)
        {
            if (Cli.TryRun(e.Args))
            {
                Shutdown();
                return;
            }
        }

        new MainWindow().Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        OfficeInterop.ShutdownApps();
        base.OnExit(e);
    }
}
