using System.Configuration;
using System.Data;
using System.IO;
using System.Windows;

namespace WebToBG
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
        public App()
        {
            AppDomain.CurrentDomain.UnhandledException += (s, e) => Log("UnhandledException", e.ExceptionObject as Exception);
            DispatcherUnhandledException += (s, e) => { Log("DispatcherUnhandledException", e.Exception); e.Handled = true; };
            TaskScheduler.UnobservedTaskException += (s, e) => { Log("UnobservedTaskException", e.Exception); e.SetObserved(); };
            System.Windows.Forms.Application.ThreadException += (s, e) => { Log("WinFormsThreadException", e.Exception); };
            
            // Handle application exit to ensure proper cleanup
            this.Exit += App_Exit;
            AppDomain.CurrentDomain.ProcessExit += (s, e) => Log("ProcessExit", null);
        }

        private void App_Exit(object sender, ExitEventArgs e)
        {
            Log("ApplicationExit", null);
            System.Diagnostics.Debug.WriteLine("[App] Application exit event triggered");
            
            // As a safety measure, force wallpaper refresh on application exit
            try
            {
                WallpaperHost.ForceWallpaperRefresh();
            }
            catch (Exception ex)
            {
                Log("WallpaperRefreshError", ex);
            }
        }

        private void Log(string kind, Exception? ex)
        {
            try
            {
                var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WebToBG");
                Directory.CreateDirectory(dir);
                File.AppendAllText(Path.Combine(dir, "debug.log"),
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {kind}: {ex}\r\n");
            }
            catch { }
        }
    }
}
