using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using Microsoft.Win32;

namespace WebToBG
{
    public static class AppTray
    {
        private static NotifyIcon? _tray;
        private static ToolStripMenuItem? _muteMenuItem;
        private static ToolStripMenuItem? _startupMenuItem;

        private const string STARTUP_REGISTRY_KEY = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private const string APP_NAME = "WebToBG";

        public static void Create(MainWindow window)
        {
            _tray = new NotifyIcon
            {
                Text = "WebToBG",
                Icon = LoadTrayIcon(),
                Visible = true
            };

            var menu = new ContextMenuStrip();

            // URL
            var setUrl = new ToolStripMenuItem("Set URL…", null, (_, __) => window.PromptAndSetUrl());

            // Actions
            var interact = new ToolStripMenuItem("Toggle Interaction (Ctrl+Alt+W)", null, (_, __) => window.ToggleInteract());
            
            // Mute with checkmark (starts checked since we start muted)
            _muteMenuItem = new ToolStripMenuItem("Muted (Ctrl+Alt+M)", null, (_, __) => 
            {
                window.ToggleMute();
                UpdateMuteMenuItem(window);
            })
            {
                Checked = true // Start checked since we start muted
            };
            
            var reload = new ToolStripMenuItem("Reload Page (Ctrl+Alt+R)", null, (_, __) => window.Reload());

            // Startup with checkmark
            _startupMenuItem = new ToolStripMenuItem("Start with Windows", null, (_, __) => 
            {
                ToggleStartup();
                UpdateStartupMenuItem();
            })
            {
                CheckOnClick = true,
                Checked = IsStartupEnabled() // Check if startup is enabled
            };

            // About submenus
            var aboutMenu = new ToolStripMenuItem("About");
            var aboutVersion = new ToolStripMenuItem($"WebToBG v{GetVersion()}", null, (_, __) => OpenGitHub());
            var aboutSeparator = new ToolStripSeparator();
            var aboutBMAC = new ToolStripMenuItem("Buy Me a Coffee", null, (_, __) => OpenBMAC());
            var aboutAuthor = new ToolStripMenuItem("Made with ❤️ by LDY681")
            {
                Enabled = false // Disable this item as it's not an action
            };

            // Quit
            var quit = new ToolStripMenuItem("Quit", null, (_, __) => System.Windows.Application.Current.Shutdown());

            // Build menu
            menu.Items.Add(setUrl);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(interact);
            menu.Items.Add(_muteMenuItem);
            menu.Items.Add(reload);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(_startupMenuItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(aboutMenu);
            aboutMenu.DropDownItems.Add(aboutVersion);
            aboutMenu.DropDownItems.Add(aboutSeparator);
            aboutMenu.DropDownItems.Add(aboutBMAC);
            aboutMenu.DropDownItems.Add(aboutAuthor);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(quit);

            _tray.ContextMenuStrip = menu;
        }

        private static string GetVersion()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var version = assembly.GetName().Version;
                return version?.ToString(3) ?? "1.0.0"; // Major.Minor.Build
            }
            catch
            {
                return "1.0.0";
            }
        }

        private static void OpenGitHub()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "https://github.com/LDY681/WebToBG",
                    UseShellExecute = true
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppTray] Failed to open GitHub link: {ex.Message}");
                MessageBox.Show($"Could not open GitHub link:\n{ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static void OpenBMAC()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "https://buymeacoffee.com/ldydyxw",
                    UseShellExecute = true
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppTray] Failed to open GitHub link: {ex.Message}");
                MessageBox.Show($"Could not open BMAC link:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static bool IsStartupEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(STARTUP_REGISTRY_KEY, false);
                return key?.GetValue(APP_NAME) != null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppTray] Error checking startup status: {ex.Message}");
                return false;
            }
        }

        private static void ToggleStartup()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(STARTUP_REGISTRY_KEY, true);
                if (key == null)
                {
                    MessageBox.Show("Unable to access Windows startup settings.", "Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (IsStartupEnabled())
                {
                    // Remove from startup
                    key.DeleteValue(APP_NAME, false);
                    System.Diagnostics.Debug.WriteLine("[AppTray] Removed from Windows startup");
                }
                else
                {
                    // Add to startup
                    var exePath = Process.GetCurrentProcess().MainModule?.FileName;
                    if (!string.IsNullOrEmpty(exePath))
                    {
                        key.SetValue(APP_NAME, $"\"{exePath}\"");
                        System.Diagnostics.Debug.WriteLine($"[AppTray] Added to Windows startup: {exePath}");
                    }
                    else
                    {
                        MessageBox.Show("Unable to determine application path.", "Error", 
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppTray] Error toggling startup: {ex.Message}");
                MessageBox.Show($"Error updating startup settings:\n{ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static Icon LoadTrayIcon()
        {
            try
            {
                // Try to load custom tray icon from embedded resources
                var assembly = Assembly.GetExecutingAssembly();
                using var stream = assembly.GetManifestResourceStream("WebToBG.Resources.tray-icon.ico");
                
                if (stream != null)
                {
                    return new Icon(stream);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppTray] Failed to load custom icon: {ex.Message}");
            }

            // Fallback to system icon if custom icon fails to load
            return SystemIcons.Application;
        }

        private static void UpdateMuteMenuItem(MainWindow window)
        {
            if (_muteMenuItem != null)
            {
                bool isMuted = window.IsMuted();
                _muteMenuItem.Checked = isMuted;
                _muteMenuItem.Text = isMuted ? "Muted (Ctrl+Alt+M)" : "Unmuted (Ctrl+Alt+M)";
            }
        }

        private static void UpdateStartupMenuItem()
        {
            if (_startupMenuItem != null)
            {
                bool isStartupEnabled = IsStartupEnabled();
                _startupMenuItem.Checked = isStartupEnabled;
            }
        }

        public static void RefreshMuteState(MainWindow window)
        {
            UpdateMuteMenuItem(window);
        }

        public static void Dispose()
        {
            if (_tray != null)
            {
                _tray.Visible = false;
                _tray.Dispose();
                _tray = null;
            }
        }
    }
}
