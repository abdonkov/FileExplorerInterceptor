# File Explorer Interceptor
 Intercept Windows File Explorer directory opening and open the directory in another file manager instead.
 
 ![fei](https://user-images.githubusercontent.com/10236674/139499324-3802f5ba-24d1-4c33-aff2-3c7826d70d9c.gif)



## Why was this made?
To provide a different solution (albeit not perfect again) for making a third party file manager application the default.

#### Background

Most file managers for Windows that are used as an alternative to Windows File Explorer have an option to be registered as the "default" file manger.

Most file managers do this by editing the registry and introducing a shell command that opens the file manager and then setting it as the default option for shell opening.

For example registering a new shell command under `HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell` and setting it as the `(Default)` value.

By doing this Windows uses this new command when the Windows Shell opens a new directory.

#### The Problem

First and foremost using this approach item selection on directory opening doesn't work, or at least not it some file managers.

However, the real problem actually is that when using this approach and using the option "Show in folder" in some apps there is a bug, which makes the directory opening very slow (more than 10 seconds) when the operation is performed a second time in a short period.

This problem probably happens only with Browsers *(Chrome, Firefox)* and Electron based apps *(Skype, Teams, Slack, Discord)* but I'm not completely sure. I had only one native app that has a "show in folder" option which uses the Windows Shell and it worked correctly there.

## How does it work?

This is an application that only has a tray icon in the system tray (mainly for easier exit if needed) and runs in the background.

- The application watches for Window Creation Events *(by setting a WinEventHook from the Win32 API)*.

- When a new file explorer window is created, the application checks what folder is opened and what files are selected *(through Shell Objects for Scripting API)*.

- And if the file explorer window is inside a folder, the window is closed and the "replacement" application is opened in the given folder.

Also, a good thing about this approach is that if an application starts the explorer.exe directly instead of using the Windows Shell *(with `explorer.exe /select,"C:\Windows\explorer.exe"` for example)*, the application replaces it with the defined file manager *(which the Registry edit for default Window Shell opening can't do, obvously)*.

*Note: File Explorer is closed only when it was opened directly inside a directory. This ensures that you can still use the Windows File Explorer normally, if you start it from "This PC", because there isn't an opened directory in that context. Also the Control Panel is working as well with this approach (yeah, I know, the Control Panel is File Explorer as well... obviously, right Microsoft?).*

## Usage

You can download the latest release from the "[Releases](https://github.com/abdonkov/FileExplorerInterceptor/releases)" tab.

After the program is run a tray icon will be visible with the option to exit the application if needed.
![image](https://user-images.githubusercontent.com/10236674/137640941-cb4d33df-8d74-4c95-9f9b-5585e124508b.png)

You will need to disable the "register as default file manager" option in your preferred file manager in order for the application to work. *(Or remove the `(Default)` value for the shell if you know what setting your file manager has changed)*.

If you want to start the application on system startup, the easiest way is to create a shortcut of the .exe and place it inside `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup`.

## Configuration

Inside the main directory of the application is a file `appsettings.json` that is used for the configuration:
````json
{
  "ApplicationPath": "%LOCALAPPDATA%\\Microsoft\\WindowsApps\\files.exe",
  "DirectoryArgumentsLine": "-Directory \"<D>\"",
  "SelectedItemArgumentsLine":  "-Directory \"<D>\" -Select \"<S>\""
}
````
*Note that backslashes (`\`) and quotes (`"`) have to be escaped with a backslash (`\`).*

`ApplicationPath` is the path to the application that will be opened as the replacement for the file explorer.

*Note: Special folder variables like `%APPDATA%` are supported only at the start of the string, just in case there is some folder with a special name (same as an environment variable) which will break the path if expanded.*

`DirectoryArgumentsLine` is the arguments line that will be used when a directory has to be opened without any selected items.

`SelectedItemArgumentsLine` is the arguments line that will be used when a directory has to be opened with a selected item inside. *Fallbacks to `DirectoryArgumentsLine` if not provided.*

The `<D>` and `<S>` are special strings that will be replaced with the ***D**irectory Path* and the ***S**elected Item* respectivly.

For example, for the given command `-Directory "<D>" -Select "<S>"` and an opened directory `C:\Windows` with selected item `explorer.exe`, the command will be transformed to `-Directory "C:\Windows" -Select "explorer.exe"`.

*Note: The default `appsettings.json` file in the application (which is the same as the above example) is configured for the ["Files" file manager](https://github.com/files-community/Files), however the directory opening with a selected item is still only available in the insider build.
When the functionallity is merged into the main build, it will work for you without having to change the configuration.*

## Limitations

1. There is a slight delay where the File Explorer is visible before it is closed, because until a newly created window is visible, it's state can't be changed (hid or minimized for example).

    * I have also tried tinkering with hooking into the explorer process and intercepting the CreateWindow call to make it start hidden or atleast minimized, but that crashes the process. Also creating a global windows WH_CBT hook doesn't allow changes to the window or the explorer crashes again. I tested a WH_SHELL hook as well for window creation, but it isn't much faster that the Window Created event.*

2. When a File Explorer is alredy opened in the given directory, Windows just focuses the existing window instead of creating a new one and thus our approach doesn't work.

    *Technically, there is a Window Focused/Activated Event that can detect that, but the user selecting the window could also fire it, which won't be great for normal usage of the File Explorer when needed.*
