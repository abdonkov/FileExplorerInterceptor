# File Explorer Interceptor
 Intercept Windows File Explorer directory opening and open the directory in another file manager instead.

## Why was this made?
To provide a different solution (albeit not perfect again) for making a third party file manager application the default.

#### Background

Most file managers for Windows that are used as an alternative to Windows File Explorer have an option to be regisred as the "default" file manger.

Most file managers do this by editing the registry and introducing a shell command that opens the file manager and then setting it as the default option for shell opening.

For example registering a new shell command under `HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell` and setting it as the `(Default)` value.

By doing this Windows uses this new command when the Windows Shell opens a new directoy.

#### The Problem

When using this approach and using the option "Show in folder" in some apps there is a bug, which makes the directory opening very slow (more than 10 seconds) when the operation is performed a second time in a short period.

This problem probably happens only with Browsers *(Chrome, Firefox)* and Electron based apps *(Skype, Teams, Slack, Discord)* but I'm not completely sure. I had only one native app that has a "show in folder" option which uses the Windows Shell and it worked correctly there.

## How does it work?

This is an application that only has a tray icon in the system tray (mainly for easier exit if needed) and runs in the background.

- The application watches for Window Creation Events *(through Windows Automation API)*.

- When a new file explorer window is created, the application checks what folder is opened and what files are selected *(through Shell Objects for Scripting API)*.

- And if the file explorer window is inside a folder, the window is closed and the "replacement" application is opened in the given folder.

Also, a good thing about this approach is that if an application starts the explorer.exe directly instead of using the Windows Shell *(with `explorer.exe /select,"C:\Windows\explorer.exe"` for example)*, the application replaces it with the defined file manager *(which the Registry edit for default Window Shell opening can't do, obvously)*.

*Note: File Explorer is closed only when it was opened directly inside a directory. This ensures that you can still use the Windows File Explorer normally, if you start it from "This PC", because there isn't an opened directory in that context. Also the Control Panel is working as well with this approach (yeah, I know, the Control Panel is File Explorer as well... obviously, right Microsoft?).*

## Configuration

Inside the main directory of the application there is a file `appsettings.json` that is used for the configuration:
````json
{
  "ApplicationPath": "%LOCALAPPDATA%\\Microsoft\\WindowsApps\\files.exe",
  "DirectoryArgumentsLine": "-Directory \"<D>\"",
  "SelectedItemArgumentsLine":  "-Directory \"<D>\" -Select \"<S>\""
}
````
`ApplicationPath` is the path to the application that will be opened as the replacement for the file explorer.

*Note: Special folder variables like `%APPDATA%` are supported only at the start of the string, just in case there is some folder with a special name (same as an environment variable) which will break the path if expanded.*

`DirectoryArgumentsLine` is the arguments line that will be used when a directory has to be opened without any selected files.

`SelectedItemArgumentsLine` is the arguments line that will be used when a directory has to be opened with a selected item inside. *Fallbacks to `DirectoryArgumentsLine` if not provided.*

The `<D>` and `<S>` are special strings that will be replaced with the **D**irectory Path and the **S**elected Item respectivly.

For example, for the given command `-Directory "<D>" -Select "<S>"` and an opened directory `C:\Windows` with selected item `explorer.exe`, the command will be transformed to `-Directory "C:\Windows" -Select "explorer.exe"`.

*Note: The default `appsettings.json` file in the application (which is the same as the above example) is configured for the "[Files](https://github.com/files-community/Files)" file manager, however the directory opening with a selected file is still not implemented for it and this was just an example created by me.
When the functionallity is implemented by the "Files" file manager I will change the default configuration file and update this documentation if needed.*





## Limitations

1. Because the Window Creation Event is not immediate, there is a slight delay where the File Explorer is visible before it is closed.

    *Technically, there is a way to hook into the window creation much earlier, but to do it globally, DLL injection will be needed which isn't a very good approach and I'm also not really familiar with that.*

2. When a File Explorer is alredy opened in the given directory, Windows just focuses the existing window instead of creating a new one and thus our approach doesn't work.

    *Technically, there is a Window Focused/Activated Event that can detect that, but the user selecting the window could also fire it, which won't be great for normal usage of the File Explorer when needed.*






