using System;
using System.Runtime.InteropServices;
using System.Windows.Forms; // WinForms
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace WebToBG
{
    internal sealed class WallpaperHost : IDisposable
    {
        // ===== P/Invoke =====
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int SW_SHOW = 5;
        private const int SW_HIDE = 0;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);
        private static readonly IntPtr HWND_TOP = new(-1);
        private static readonly IntPtr HWND_BOTTOM = new(1);
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_SHOWWINDOW = 0x0040;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll", EntryPoint = "GetWindowLong", SetLastError = true)]
        private static extern IntPtr GetWindowLong32(IntPtr hWnd, int nIndex);
        private static IntPtr GetWindowLongAuto(IntPtr hWnd, int nIndex)
            => IntPtr.Size == 8 ? GetWindowLongPtr64(hWnd, nIndex) : GetWindowLong32(hWnd, nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);
        [DllImport("user32.dll", EntryPoint = "SetWindowLong", SetLastError = true)]
        private static extern IntPtr SetWindowLong32(IntPtr hWnd, int nIndex, IntPtr dwNewLong);
        private static IntPtr SetWindowLongAuto(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
            => IntPtr.Size == 8 ? SetWindowLongPtr64(hWnd, nIndex, dwNewLong)
                                : SetWindowLong32(hWnd, nIndex, dwNewLong);

        private const int GWL_EXSTYLE = -20;
        private const uint WS_EX_TRANSPARENT = 0x00000020;
        private const uint WS_EX_NOACTIVATE = 0x08000000;

        // ===== Wallpaper restoration P/Invoke =====
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool SystemParametersInfo(uint uiAction, uint uiParam, string pvParam, uint fWinIni);
        
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool InvalidateRect(IntPtr hWnd, IntPtr lpRect, bool bErase);
        
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UpdateWindow(IntPtr hWnd);
        
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetDesktopWindow();
        
        private const uint SPI_SETDESKWALLPAPER = 0x0014;
        private const uint SPIF_UPDATEINIFILE = 0x01;
        private const uint SPIF_SENDCHANGE = 0x02;

        /// <summary>
        /// Static method to force Windows to refresh/restore the original desktop wallpaper.
        /// </summary>
        public static void ForceWallpaperRefresh()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("[WallpaperHost] Force refreshing desktop wallpaper...");
                
                // Force Windows to refresh the desktop wallpaper
                SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, string.Empty, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);
                
                // Additional desktop refresh
                var desktopWindow = GetDesktopWindow();
                InvalidateRect(desktopWindow, IntPtr.Zero, true);
                UpdateWindow(desktopWindow);
                
                System.Diagnostics.Debug.WriteLine("[WallpaperHost] Force wallpaper refresh complete");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WallpaperHost] ForceWallpaperRefresh error: {ex}");
            }
        }

        // ===== Single Form & WebView =====
        private Form _webForm = null!;
        private WebView2 _webView = null!;
        private CoreWebView2Environment _env = null!;
        private IntPtr _workerW = IntPtr.Zero; // Store WorkerW handle
        private bool _isDisposed = false;

        private string _currentUrl = "https://www.google.com";
        private bool _isMuted = true; // Always start muted
        private bool _isInteractive = false; // Start non-interactive

        public CoreWebView2? Core => _webView?.CoreWebView2;
        public bool IsMuted => _isMuted;

        // ===== UI thread marshaller =====
        private void OnForm(Action action)
        {
            if (_isDisposed) return;
            if (_webForm?.IsHandleCreated == true && _webForm.InvokeRequired) 
                _webForm.Invoke(action);
            else 
                action();
        }

        public async System.Threading.Tasks.Task InitAsync(string userDataPath, string url, bool startMuted)
        {
            _currentUrl = url;
            _isMuted = true; // Always start muted

            // Create shared environment with audio permissions
            var options = new CoreWebView2EnvironmentOptions
            {
                AdditionalBrowserArguments = "--autoplay-policy=no-user-gesture-required --enable-features=AudioServiceOutOfProcess"
            };
            
            _env = await CoreWebView2Environment.CreateAsync(
                browserExecutableFolder: null,
                userDataFolder: userDataPath,
                options: options);

            // Create form as full-screen wallpaper
            _webForm = new Form
            {
                Text = "WebToBG",
                FormBorderStyle = FormBorderStyle.None,
                StartPosition = FormStartPosition.Manual,
                ShowInTaskbar = false,
                TopMost = false,
                Bounds = Screen.PrimaryScreen?.Bounds ?? new System.Drawing.Rectangle(0, 0, 1920, 1080),
                Icon = LoadApplicationIcon() // Set window icon
            };

            _webView = new WebView2 { Dock = DockStyle.Fill };
            _webForm.Controls.Add(_webView);
            
            // Show form and initialize WebView
            _webForm.Show();
            await _webView.EnsureCoreWebView2Async(_env);
            
            // Configure WebView2 settings for audio
            var settings = _webView.CoreWebView2.Settings;
            settings.AreDevToolsEnabled = false;
            settings.AreDefaultContextMenusEnabled = true;
            settings.AreHostObjectsAllowed = true;
            settings.IsWebMessageEnabled = true;
            settings.IsGeneralAutofillEnabled = true;
            
            // Enable audio and set initial mute state
            _webView.CoreWebView2.IsMuted = true; // Start muted by default
            
            // Add permission handler for audio
            _webView.CoreWebView2.PermissionRequested += (sender, args) =>
            {
                if (args.PermissionKind == CoreWebView2PermissionKind.Microphone ||
                    args.PermissionKind == CoreWebView2PermissionKind.Camera ||
                    args.PermissionKind == CoreWebView2PermissionKind.Geolocation)
                {
                    args.State = CoreWebView2PermissionState.Deny;
                }
                else
                {
                    args.State = CoreWebView2PermissionState.Allow;
                }
            };
            
            _webView.Source = new Uri(_currentUrl);
            _webView.CoreWebView2.ProcessFailed += (s, e) =>
            {
                System.Diagnostics.Debug.WriteLine($"[WebView2] ProcessFailed: {e.ProcessFailedKind}");
            };

            // Start as wallpaper (non-interactive)
            _isInteractive = false;
        }

        private static System.Drawing.Icon LoadApplicationIcon()
        {
            try
            {
                // Try to load custom application icon from embedded resources
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using var stream = assembly.GetManifestResourceStream("WebToBG.Resources.app-icon.ico");
                
                if (stream != null)
                {
                    return new System.Drawing.Icon(stream);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WallpaperHost] Failed to load custom icon: {ex.Message}");
            }

            // Fallback to system icon if custom icon fails to load
            return System.Drawing.SystemIcons.Application;
        }

        public void Navigate(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return;
            if (!url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                !url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) &&
                !url.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
            {
                url = "https://" + url.Trim();
            }
            _currentUrl = url;

            OnForm(() => _webView.Source = new Uri(_currentUrl));
        }

        public void Reload()
        {
            OnForm(() => _webView.Reload());
        }

        public void SetMuted(bool value)
        {
            _isMuted = value;
            OnForm(() => 
            { 
                if (_webView.CoreWebView2 != null) 
                {
                    _webView.CoreWebView2.IsMuted = value;
                    System.Diagnostics.Debug.WriteLine($"[WallpaperHost] WebView2.IsMuted set to: {value}");
                    System.Diagnostics.Debug.WriteLine($"[WallpaperHost] WebView2.IsMuted actual value: {_webView.CoreWebView2.IsMuted}");
                }
            });
        }

        public void ToggleMute()
        {
            SetMuted(!_isMuted);
            System.Diagnostics.Debug.WriteLine($"[WallpaperHost] Mute toggled to: {_isMuted}");
        }

        public void ToBackground(IntPtr workerW)
        {
            OnForm(() =>
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine("[WallpaperHost] Setting up as wallpaper...");

                    // Store WorkerW handle for later use
                    _workerW = workerW;
                    
                    // Ensure we're not interactive first
                    _isInteractive = false;
                    
                    // Parent to WorkerW to make it a wallpaper
                    SetParent(_webForm.Handle, workerW);

                    // Size to full screen
                    var bounds = Screen.PrimaryScreen?.Bounds ?? new System.Drawing.Rectangle(0, 0, 1920, 1080);
                    SetWindowPos(_webForm.Handle, IntPtr.Zero, 
                        bounds.X, bounds.Y, bounds.Width, bounds.Height,
                        SWP_NOACTIVATE | SWP_NOZORDER | SWP_SHOWWINDOW);

                    // Set non-interactive
                    SetInteractive(false);
                    
                    System.Diagnostics.Debug.WriteLine("[WallpaperHost] Wallpaper setup complete");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[WallpaperHost] ToBackground error: {ex}");
                }
            });
        }

        public void ToggleInteractive()
        {
            OnForm(() =>
            {
                try
                {
                    _isInteractive = !_isInteractive;
                    SetInteractive(_isInteractive);
                    
                    System.Diagnostics.Debug.WriteLine($"[WallpaperHost] Interaction {(_isInteractive ? "enabled" : "disabled")}");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[WallpaperHost] ToggleInteractive error: {ex}");
                }
            });
        }

        private void SetInteractive(bool interactive)
        {
            OnForm(() =>
            {
                var h = _webForm.Handle;
                if (h == IntPtr.Zero) return;
                
                System.Diagnostics.Debug.WriteLine($"[WallpaperHost] Setting interactive mode: {interactive}");
                
                long style = GetWindowLongAuto(h, GWL_EXSTYLE).ToInt64();
                System.Diagnostics.Debug.WriteLine($"[WallpaperHost] Current window style: 0x{style:X}");
                
                if (interactive)
                {
                    // Make interactive: remove transparent and noactivate flags
                    style &= ~((int)WS_EX_TRANSPARENT);
                    style &= ~((int)WS_EX_NOACTIVATE);
                    
                    System.Diagnostics.Debug.WriteLine($"[WallpaperHost] New interactive style: 0x{style:X}");
                    SetWindowLongAuto(h, GWL_EXSTYLE, new IntPtr(style));
                    
                    // Remove from WorkerW and make it a normal window
                    SetParent(h, IntPtr.Zero);
                    
                    // Bring to front when interactive
                    _webForm.TopMost = true;
                    _webForm.Show();
                    _webForm.Activate();
                    _webForm.Focus();
                    
                    SetWindowPos(h, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                    
                    System.Diagnostics.Debug.WriteLine("[WallpaperHost] Window brought to front for interaction");
                }
                else
                {
                    // Make non-interactive: add transparent and noactivate flags
                    style |= (int)WS_EX_TRANSPARENT;
                    style |= (int)WS_EX_NOACTIVATE;
                    
                    System.Diagnostics.Debug.WriteLine($"[WallpaperHost] New non-interactive style: 0x{style:X}");
                    SetWindowLongAuto(h, GWL_EXSTYLE, new IntPtr(style));
                    
                    // Re-parent to WorkerW when going back to wallpaper mode
                    if (_workerW != IntPtr.Zero)
                    {
                        SetParent(h, _workerW);
                        System.Diagnostics.Debug.WriteLine("[WallpaperHost] Re-parented to WorkerW");
                    }
                    
                    // Send to back when non-interactive
                    _webForm.TopMost = false;
                    SetWindowPos(h, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                    
                    System.Diagnostics.Debug.WriteLine("[WallpaperHost] Window sent to back");
                }
                
                // Verify the style was set
                long newStyle = GetWindowLongAuto(h, GWL_EXSTYLE).ToInt64();
                System.Diagnostics.Debug.WriteLine($"[WallpaperHost] Final window style: 0x{newStyle:X}");
                System.Diagnostics.Debug.WriteLine($"[WallpaperHost] Interactive mode set to: {interactive}");
            });
        }

        // Legacy methods for compatibility
        public void ToForeground() => SetInteractive(true);
        public void ApplyClickThrough(bool enable) => SetInteractive(!enable);
        public bool IsInBackground => !_isInteractive;

        private void RestoreOriginalWallpaper()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("[WallpaperHost] Restoring original wallpaper...");
                
                // First, remove our form from WorkerW if it's parented there
                if (_webForm != null && _webForm.Handle != IntPtr.Zero && _workerW != IntPtr.Zero)
                {
                    SetParent(_webForm.Handle, IntPtr.Zero);
                    ShowWindow(_webForm.Handle, SW_HIDE);
                }
                
                // Force Windows to refresh the desktop wallpaper
                // This will restore the original wallpaper that was set before our app started
                SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, string.Empty, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);
                
                // Additional desktop refresh
                var desktopWindow = GetDesktopWindow();
                InvalidateRect(desktopWindow, IntPtr.Zero, true);
                UpdateWindow(desktopWindow);
                
                // Also invalidate the WorkerW window specifically
                if (_workerW != IntPtr.Zero)
                {
                    InvalidateRect(_workerW, IntPtr.Zero, true);
                    UpdateWindow(_workerW);
                }
                
                System.Diagnostics.Debug.WriteLine("[WallpaperHost] Wallpaper restoration complete");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WallpaperHost] RestoreOriginalWallpaper error: {ex}");
            }
        }

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;
            
            try
            {
                System.Diagnostics.Debug.WriteLine("[WallpaperHost] Starting disposal...");
                
                // First restore the original wallpaper before disposing our form
                RestoreOriginalWallpaper();
                
                // Now dispose our resources
                _webForm?.Close();
                _webForm?.Dispose();
                
                System.Diagnostics.Debug.WriteLine("[WallpaperHost] Disposal complete");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WallpaperHost] Dispose error: {ex}");
            }
        }
    }
}
