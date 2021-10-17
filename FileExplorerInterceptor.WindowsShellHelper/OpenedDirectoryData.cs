using System;
using System.Collections.Generic;
using System.Text;

namespace FileExplorerInterceptor.WindowsShellHelper
{
    /// <summary>
    /// Represents the data of an opened directory inside windows explorer.
    /// </summary>
    public class OpenedDirectoryData
    {
        /// <summary>
        /// The path of the opened directory
        /// </summary>
        public string Path { get; set; }

        /// <summary>
        /// A list of the selected items names.
        /// </summary>
        public IList<string> SelectedItems { get; set; }
    }
}
