using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace FileExplorerInterceptor.WindowsShell
{
    public static class ShellReader
    {
        public static OpenedDirectoryData CloseFileExplorerIfDirectoryOpenedAndGetDirectoryPath(long windowHandle, bool withSelectedItems = false)
        {
            return GetFileExplorerOpenedDirectoryDataIfAny(windowHandle, withSelectedItems,
            (item, itemType) =>
            {
                try
                {
                    itemType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, item, null);
                }
                catch { }
            });
        }

        public static OpenedDirectoryData GetFileExplorerOpenedDirectoryDataIfAny(long windowHandle, bool withSelectedItems = false, Action<object, Type> onFountItemAction = null)
        {
            try
            {
                Type shellType = Type.GetTypeFromProgID("Shell.Application");
                object shell = Activator.CreateInstance(shellType);

                object windows = shellType.InvokeMember("Windows", BindingFlags.InvokeMethod, null, shell, new object[] { });
                Type windowsType = windows.GetType();
                int windowsCount = (int)windowsType.InvokeMember("Count", BindingFlags.GetProperty, null, windows, null);

                for (int i = 0; i < windowsCount; i++)
                {
                    object item = windowsType.InvokeMember("Item", BindingFlags.InvokeMethod, null, windows, new object[] { i });
                    Type itemType = item.GetType();

                    long itemHwnd = (long)itemType.InvokeMember("HWND", BindingFlags.GetProperty, null, item, null);
                    if (itemHwnd != windowHandle)
                        continue;

                    var itemName = itemType.InvokeMember("Name", BindingFlags.GetProperty, null, item, null) as string;
                    var itemFullName = itemType.InvokeMember("FullName", BindingFlags.GetProperty, null, item, null) as string;

                    if (string.Equals(itemName, "Windows Explorer", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(itemName, "File Explorer", StringComparison.OrdinalIgnoreCase)
                        || (itemFullName?.EndsWith("explorer.exe", StringComparison.OrdinalIgnoreCase) ?? false))
                    {
                        string locationUrl = (string)itemType.InvokeMember("LocationURL", BindingFlags.GetProperty, null, item, null);
                        if (!string.IsNullOrWhiteSpace(locationUrl))
                        {
                            if (Uri.TryCreate(locationUrl, UriKind.Absolute, out Uri uri))
                            {
                                if (uri.IsFile)
                                {
                                    var openedDirectoryData = new OpenedDirectoryData
                                    {
                                        Path = uri.LocalPath,
                                        SelectedItems = new List<string>(),
                                    };

                                    if (withSelectedItems)
                                    {
                                        // Read selected items
                                        try
                                        {
                                            object document = itemType.InvokeMember("Document", BindingFlags.GetProperty, null, item, null);
                                            Type documentType = document.GetType();

                                            object selectedItems = documentType.InvokeMember("SelectedItems", BindingFlags.InvokeMethod, null, document, null);
                                            Type selectedItemsType = selectedItems.GetType();

                                            int selectedItemsCount = (int)selectedItemsType.InvokeMember("Count", BindingFlags.GetProperty, null, selectedItems, null);
                                            for (int j = 0; j < selectedItemsCount; j++)
                                            {
                                                object selectedItem = selectedItemsType.InvokeMember("Item", BindingFlags.InvokeMethod, null, selectedItems, new object[] { j });
                                                Type selectedItemType = selectedItem.GetType();

                                                string selectedItemName = selectedItemType.InvokeMember("Name", BindingFlags.GetProperty, null, selectedItem, null) as string;
                                                if (!string.IsNullOrWhiteSpace(selectedItemName))
                                                {
                                                    openedDirectoryData.SelectedItems.Add(selectedItemName);
                                                }
                                            }
                                        }
                                        catch { }
                                    }

                                    if (onFountItemAction != null)
                                        onFountItemAction(item, itemType);

                                    return openedDirectoryData;
                                }
                            }
                        }
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }
}
