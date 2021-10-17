using System;
using System.Windows.Forms;
using System.Threading;

namespace FileExplorerInterceptor
{
    static class Program
    {
        // Register mutex to use for single instace check
        static Mutex mutex = new Mutex(false, "FileExplorerInterceptor_SingleInstance_Mutex");

        [STAThread]
        static void Main()
        {
            // Using the mutex to check if another instance of the application is already running and end this one
            if (!mutex.WaitOne(100, false)) // waits 100 ms before check just in case
            {
                return;
            }

            try
            {

#if NET5_0_OR_GREATER
                Application.SetHighDpiMode(HighDpiMode.SystemAware);
#endif
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new TrayIconApplicationContext());
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }
    }
}
