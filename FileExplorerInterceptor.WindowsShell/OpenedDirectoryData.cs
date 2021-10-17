using System;
using System.Collections.Generic;
using System.Text;

namespace FileExplorerInterceptor.WindowsShell
{
    public class OpenedDirectoryData
    {
        public string Path { get; set; }

        public IList<string> SelectedItems { get; set; }
    }
}
