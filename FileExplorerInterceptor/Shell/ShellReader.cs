using FileExplorerInterceptor.Interop;
using FileExplorerInterceptor.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Threading;

namespace FileExplorerInterceptor.Shell
{
    public static class ShellReader
    {
        public static OpenedDirectoryData CloseFileExplorerIfDirectoryOpenedAndGetDirectoryPath(long windowHandle, bool withSelectedItems, out bool hasFoundWindow, bool searchUntilFound, TimeSpan maxSearchTime, TimeSpan maxSearchTimeForSelectedItems, int waitBeforeNextRetryInMs = 30)
        {
            return GetFileExplorerOpenedDirectoryDataIfAny(windowHandle, withSelectedItems, out hasFoundWindow, searchUntilFound, maxSearchTime, maxSearchTimeForSelectedItems, waitBeforeNextRetryInMs,
            (item, itemType) =>
            {
                try
                {
                    itemType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, item, null);
                }
                catch { }
            });
        }

        public static OpenedDirectoryData GetFileExplorerOpenedDirectoryDataIfAny(long windowHandle, bool withSelectedItems, out bool hasFoundWindow, bool searchUntilFound, TimeSpan maxSearchTime, TimeSpan maxSearchTimeForSelectedItems, int waitBeforeNextRetryInMs = 30, Action<object, Type> onFoundItemAction = null)
        {
            hasFoundWindow = false;

            try
            {
                // Shell reading is done using the Shell Objects For Scripting API
                // Reflection is used intead of COM reference, because the COM reference wasn't complete and reflection was still needed for accessing some properties

                Type shellType = Type.GetTypeFromProgID("Shell.Application");
                object shell = Activator.CreateInstance(shellType);

                long maxSearchTimeTicks = maxSearchTime.Ticks;
                var stopwatch = new Stopwatch();
                stopwatch.Start();
                do
                {
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

                        hasFoundWindow = true;

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
                                            long elapsedTicksOnSelectedItemsSearchStart = stopwatch.ElapsedTicks;
                                            long remainingTicks = maxSearchTimeTicks - elapsedTicksOnSelectedItemsSearchStart;
                                            long maxSearchTimeForSelectedItemsTicks = maxSearchTimeForSelectedItems.Ticks < remainingTicks ? maxSearchTimeForSelectedItems.Ticks : remainingTicks;
                                            bool selectedItemsFound = false;

                                            do
                                            {
                                                try
                                                {
                                                    object document = itemType.InvokeMember("Document", BindingFlags.GetProperty, null, item, null);
                                                    Type documentType = document.GetType();

                                                    object selectedItems = documentType.InvokeMember("SelectedItems", BindingFlags.InvokeMethod, null, document, null);
                                                    Type selectedItemsType = selectedItems.GetType();

                                                    int selectedItemsCount = (int)selectedItemsType.InvokeMember("Count", BindingFlags.GetProperty, null, selectedItems, null);
                                                    for (int j = 0; j < selectedItemsCount; j++)
                                                    {
                                                        selectedItemsFound = true;

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
                                            while (!selectedItemsFound && stopwatch.ElapsedTicks - elapsedTicksOnSelectedItemsSearchStart <= maxSearchTimeForSelectedItemsTicks);


                                        }

                                        if (onFoundItemAction != null)
                                            onFoundItemAction(item, itemType);

                                        return openedDirectoryData;
                                    }
                                }
                            }
                        }
                    }

                    Thread.Sleep(waitBeforeNextRetryInMs);
                }
                while (searchUntilFound && !hasFoundWindow && stopwatch.ElapsedTicks <= maxSearchTimeTicks);
                stopwatch.Stop();

                return null;
            }
            catch (Exception ex)
            {
                return null;
            }

        }
    }
}
