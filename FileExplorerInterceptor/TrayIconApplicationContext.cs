using FileExplorerInterceptor.Properties;
using FileExplorerInterceptor.WindowsShellHelper;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Primitives;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Automation;
using System.Windows.Forms;

namespace FileExplorerInterceptor
{
    public class TrayIconApplicationContext : ApplicationContext
    {
        private NotifyIcon trayIcon;

        private IConfiguration configuration;
        private AppSettings config;

        private Regex specialFolderRegex = new Regex("(^%.*?%)", RegexOptions.Compiled);
        private Regex directoryArgumentRegex = new Regex("<(d|D)>");
        private Regex selectedItemArgumentRegex = new Regex("<(s|S)>");

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
            Automation.AddAutomationEventHandler(
                WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement,
                TreeScope.Children,
                OnWindowOpenedHandler);
        }

        private void OnWindowOpenedHandler(object sender, AutomationEventArgs e)
        {
            try
            {
                var element = sender as AutomationElement;
                if (element == null) return;

                if (element.Current.ClassName == "CabinetWClass") // Class of all file explorer windows
                {
                    var process = Process.GetProcessById(element.Current.ProcessId);
                    if (process.ProcessName == "explorer") // check if the process is explorer, because other windows can use the same class
                    {
                        // Try to get opened directory data and close the explorer windows if inside a directory
                        var openedDirectoryData = ShellReader.CloseFileExplorerIfDirectoryOpenedAndGetDirectoryPath(element.Current.NativeWindowHandle, true);
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

            Automation.RemoveAllEventHandlers();

            Application.Exit();
        }
    }
}
