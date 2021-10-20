using FileExplorerInterceptor.Interop;
using FileExplorerInterceptor.Interop.Delegates;
using FileExplorerInterceptor.Models;
using FileExplorerInterceptor.Properties;
using FileExplorerInterceptor.Shell;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Primitives;
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace FileExplorerInterceptor
{
    public class TrayIconApplicationContext : ApplicationContext
    {
        private NotifyIcon trayIcon;

        private IConfiguration configuration;
        private AppSettings config;

        private WinEventDelegate winEventDelegate = null;
        private IntPtr winEventHook;

        private Regex specialFolderRegex = new Regex("(^%.*?%)", RegexOptions.Compiled);
        private Regex directoryArgumentRegex = new Regex("<(d|D)>", RegexOptions.Compiled);
        private Regex selectedItemArgumentRegex = new Regex("<(s|S)>", RegexOptions.Compiled);

        public TrayIconApplicationContext()
        {
            var contextMenu = new ContextMenuStrip();

            contextMenu.Items.Add("Exit", null, Exit);

            // Create tray icon
            trayIcon = new NotifyIcon
            {
                Icon = Resources.AppIcon,
                Text = "File Explorer Interceptor",
                ContextMenuStrip = contextMenu,
                Visible = true
            };

            // Build Configuration
            configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            config = new AppSettings();
            configuration.Bind(config);

            ChangeToken.OnChange(() => configuration.GetReloadToken(), () => configuration.Bind(config));

            // Register window open handler
            winEventDelegate = new WinEventDelegate(OnWindowOpenHandler);
            winEventHook = User32.SetWinEventHook(User32.EVENT_OBJECT_CREATE, User32.EVENT_OBJECT_CREATE, IntPtr.Zero, winEventDelegate, 0, 0, User32.WINEVENT_OUTOFCONTEXT);
        }

        private void OnWindowOpenHandler(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            try
            {
                var windowClass = new StringBuilder(255);
                User32.RealGetWindowClass(hwnd, windowClass, 255);

                if (windowClass.ToString() == "CabinetWClass")
                {
                    User32.GetWindowThreadProcessId(hwnd, out IntPtr pid);
                    var process = Process.GetProcessById(pid.ToInt32());
                    if (process.ProcessName == "explorer")
                    {
                        // Try to get opened directory data and close the explorer windows if inside a directory
                        var openedDirectoryData = ShellReader.CloseFileExplorerIfDirectoryOpenedAndGetDirectoryPath(
                            hwnd.ToInt32(), withSelectedItems: true, out bool foundWindow, searchUntilFound: true, maxSearchTime: TimeSpan.FromSeconds(1));

                        if (openedDirectoryData != null
                            && !string.IsNullOrWhiteSpace(openedDirectoryData.Path))
                        {
                            OpenApplicationWithDirecory(openedDirectoryData);
                        }
                    }
                }
            }
            catch { }
        }

        private void OpenApplicationWithDirecory(OpenedDirectoryData openedDirectoryData)
        {
            var applicationPath = config.ApplicationPath;
            var directoryArgumentsLine = config.DirectoryArgumentsLine;
            var selectedItemArgumentsLine = config.SelectedItemArgumentsLine;

            if (string.IsNullOrWhiteSpace(selectedItemArgumentsLine))
                selectedItemArgumentsLine = directoryArgumentsLine; // falback to directory arguments line when the selected item one is not defined

            if (string.IsNullOrWhiteSpace(applicationPath))
            {
                MessageBox.Show("Application path is not defined in the configuration!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string fileName = applicationPath;
            string arguments = null;

            // Check if application path uses a special folder and expand the path
            fileName = specialFolderRegex.Replace(fileName, (match) => Environment.ExpandEnvironmentVariables(match.Value));

            // Replace arguments list with parameters
            if (openedDirectoryData.SelectedItems?.Count > 0
                && !string.IsNullOrWhiteSpace(openedDirectoryData.SelectedItems[0])
                && !string.IsNullOrWhiteSpace(selectedItemArgumentsLine))
            {
                arguments = directoryArgumentRegex.Replace(selectedItemArgumentsLine, openedDirectoryData.Path);
                arguments = selectedItemArgumentRegex.Replace(arguments, openedDirectoryData.SelectedItems[0]);
            }
            else if (!string.IsNullOrWhiteSpace(directoryArgumentsLine))
            {
                arguments = directoryArgumentRegex.Replace(directoryArgumentsLine, openedDirectoryData.Path);
            }

            try
            {
                // Start new process
                var processStartInfo = new ProcessStartInfo();
                processStartInfo.FileName = Path.GetFullPath(fileName);
                processStartInfo.Arguments = !string.IsNullOrWhiteSpace(arguments) ? arguments : string.Empty;
                processStartInfo.UseShellExecute = true; // Runs as separate process (not a child of this process)
                Process.Start(processStartInfo);
            }
            catch
            {
                MessageBox.Show("Could not start the process! Check your configaration file and ensure correct path to application!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Exit(object sender, EventArgs e)
        {
            trayIcon.Visible = false;

            User32.UnhookWinEvent(winEventHook);

            Application.Exit();
        }
    }
}
