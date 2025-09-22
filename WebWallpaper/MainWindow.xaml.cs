using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using Microsoft.VisualBasic;
using Microsoft.Web.WebView2.Core;

namespace WebToBG
{
    public partial class MainWindow : Window
    {
        // ========= Settings =========
        private class AppSettings
        {
            public string? Url { get; set; } = "https://www.google.com";
        }

        private AppSettings _settings = new();
        private string SettingsDir =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WebToBG");
        private string SettingsPath => Path.Combine(SettingsDir, "settings.json");
        private const string USER_DATA_DIR_NAME = "WebToBGUserData";

        private void LoadSettings()
        {
            try
            {
                Directory.CreateDirectory(SettingsDir);
                if (File.Exists(SettingsPath))
                {
                    var json = File.ReadAllText(SettingsPath);
                    _settings = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                }
            }
            catch { _settings = new AppSettings(); }
        }
        private void SaveSettings()
        {
            try
            {
                Directory.CreateDirectory(SettingsDir);
                var json = JsonSerializer.Serialize(_settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsPath, json);
            }
            catch { }
        }

        // ========= Win32 (WorkerW find + hotkeys) =========
        private const uint WM_HOTKEY = 0x0312;
        private const uint WM_SPAWN_WORKERW = 0x052C;
        private const uint SMTO_NORMAL = 0x0000;

        private const int MOD_ALT = 0x0001;
        private const int MOD_CONTROL = 0x0002;
        private const int HOTKEY_ID_TOGGLE = 0x1001;
        private const int HOTKEY_ID_MUTE = 0x1002;
        private const int HOTKEY_ID_RELOAD = 0x1003;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string? lpWindowName);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string? className, string? windowTitle);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam,
            uint flags, uint timeout, out IntPtr lpdwResult);

        // ========= Fields =========
        private IntPtr _hwnd = IntPtr.Zero;         // handle of this WPF window (used only for hotkeys)
        private IntPtr _workerW = IntPtr.Zero;      // wallpaper host

#pragma warning disable IDE0044 // Add readonly modifier
        private WallpaperHost _host = new();        // Single WebView2 host
#pragma warning restore IDE0044 // Add readonly modifier

        public MainWindow()
        {
            InitializeComponent();
        }

        // ========= Lifecycle =========
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // We can keep the WPF window hidden entirely; the visible window is the WinForms host.
            this.Visibility = Visibility.Hidden;

            LoadSettings();

            var userDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                USER_DATA_DIR_NAME);

            await _host.InitAsync(userDataPath, _settings.Url ?? "https://www.google.com", startMuted: true);

            // Install hotkeys on this window
            var helper = new WindowInteropHelper(this);
            _hwnd = helper.EnsureHandle();
            var source = HwndSource.FromHwnd(_hwnd);
            source.AddHook(WndProc);

            RegisterHotKey(_hwnd, HOTKEY_ID_TOGGLE, MOD_CONTROL | MOD_ALT, KeyInterop.VirtualKeyFromKey(Key.W));
            RegisterHotKey(_hwnd, HOTKEY_ID_MUTE, MOD_CONTROL | MOD_ALT, KeyInterop.VirtualKeyFromKey(Key.M));
            RegisterHotKey(_hwnd, HOTKEY_ID_RELOAD, MOD_CONTROL | MOD_ALT, KeyInterop.VirtualKeyFromKey(Key.R));

            // Tray
            AppTray.Create(this);

            // Start in wallpaper mode
            ToBackground();
        }

        protected override void OnClosed(EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[MainWindow] Application closing - starting cleanup...");

            try
            {
                // Unregister hotkeys first
                if (_hwnd != IntPtr.Zero)
                {
                    UnregisterHotKey(_hwnd, HOTKEY_ID_TOGGLE);
                    UnregisterHotKey(_hwnd, HOTKEY_ID_MUTE);
                    UnregisterHotKey(_hwnd, HOTKEY_ID_RELOAD);
                }

                // Dispose the wallpaper host (this will restore the original wallpaper)
                _host?.Dispose();

                // Dispose the tray
                AppTray.Dispose();

                System.Diagnostics.Debug.WriteLine("[MainWindow] Cleanup complete");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Error during cleanup: {ex}");
            }
            finally
            {
                base.OnClosed(e);
            }
        }

        // ========= WorkerW discovery =========
        private IntPtr FindWallpaperWorkerW()
        {
            var progman = FindWindow("Progman", null);
            SendMessageTimeout(progman, WM_SPAWN_WORKERW, IntPtr.Zero, IntPtr.Zero, SMTO_NORMAL, 1000, out _);

            IntPtr candidate = IntPtr.Zero;
            IntPtr after = IntPtr.Zero;
            while (true)
            {
                IntPtr w = FindWindowEx(IntPtr.Zero, after, "WorkerW", null);
                if (w == IntPtr.Zero) break;
                after = w;

                bool hasIcons = FindWindowEx(w, IntPtr.Zero, "SHELLDLL_DefView", null) != IntPtr.Zero;
                if (!hasIcons)
                    candidate = w; // keep last non-icons WorkerW
            }
            return candidate;
        }

        // ========= Public API used by tray =========
        public void ToggleInteract()
        {
            _host.ToggleInteractive();
        }

        public void ToggleMute()
        {
            _host.ToggleMute();
            // Refresh the tray menu to show updated mute state
            AppTray.RefreshMuteState(this);
        }

        public void Reload() => _host.Reload();

        public void PromptAndSetUrl()
        {
            var current = _settings.Url ?? "";
            var input = Interaction.InputBox("Enter URL (http/https/file):", "Set URL", current);
            if (!string.IsNullOrWhiteSpace(input))
            {
                NavigateTo(input, persist: true);
            }
        }

        public void NavigateTo(string url, bool persist = false)
        {
            try
            {
                string u = url.Trim();
                if (!u.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                    !u.StartsWith("https://", StringComparison.OrdinalIgnoreCase) &&
                    !u.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
                {
                    u = "https://" + u;
                }

                _host.Navigate(u);

                if (persist)
                {
                    _settings.Url = u;
                    SaveSettings();
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Invalid URL:\n{ex.Message}", "WebToBG",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        public bool IsMuted()
        {
            return _host.IsMuted;
        }

        // ========= Modes =========
        private void ToBackground()
        {
            _workerW = FindWallpaperWorkerW();
            if (_workerW == IntPtr.Zero)
            {
                System.Diagnostics.Debug.WriteLine("No WorkerW found; staying foreground.");
                ToForeground();
                return;
            }

            _host.ToBackground(_workerW);
        }

        private void ToForeground()
        {
            _host.ToForeground();
        }

        // ========= Hotkeys =========
        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_HOTKEY)
            {
                int id = wParam.ToInt32();
                if (id == HOTKEY_ID_TOGGLE) { ToggleInteract(); handled = true; }
                else if (id == HOTKEY_ID_MUTE) { ToggleMute(); handled = true; }
                else if (id == HOTKEY_ID_RELOAD) { Reload(); handled = true; }
            }
            return IntPtr.Zero;
        }
    }
}