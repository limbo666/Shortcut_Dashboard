# Shortcut Dashboard
Shortcut Dashboard is a PowerShell script that creates a customizable, user-friendly GUI for launching various types of files and scripts from a centralized dashboard. It detects automatically executables, scripts, shortcuts and displays a form to use as a launchpad.

## Features

- **Dynamic Button Creation**: Automatically generates buttons for executable files, batch scripts, VBS scripts, CMD files, PowerShell scripts, and shortcuts found in the 'Scripts' folder.
- **Custom Icons**: Displays appropriate icons for different file types to enhance visual recognition.
- **Tooltips**: Shows full file paths when hovering over buttons.
- **Configurable**: Uses an INI file for storing settings like form position and logging preferences.
- **Logging**: Optional logging functionality to track user actions and script operations.
- **Adaptive Layout**: Adjusts form size based on the number of buttons and screen size.
- **Shortcut Support**: Handles .lnk files by resolving and executing their targets.

## Setup

1. Extract the files from release package to your desired folder.
3. Create a 'Scripts' folder in the same directory (you can use the demo folder and it contents for testing).
4. Add your executable files, scripts, and shortcuts to the 'Scripts' folder.

## Usage

1. Run the `LaunchShortcutDashboard1.1.vbs` script or you can run `ShortcutDashboard1.7.ps1` if you are ok with the powershell window to be visible on your working area.
2. The dashboard will appear with buttons for each file in the 'Scripts' folder.
3. Click on a button to execute the corresponding file or script.

## Configuration

- `options.ini`: Contains settings for form position and logging.
- Logging: Set `LogEnabled=1` in the INI file to enable logging to `ShortcutDashboard_Log.txt`.

## Requirements

- Windows operating system
- PowerShell 5.1 or later

## Notes

- The script automatically saves and loads the form position between sessions.
- The files and shortcuts on Scripts folder can be organized in subfolders as the script is able to scan subfodlers as well 

  
### Preview screens
![image](https://github.com/user-attachments/assets/e8ece3dc-1382-4f40-86c0-d25a82652671) ![image](https://github.com/user-attachments/assets/d213a2b0-7492-4d94-9009-99c21e57c3e9) ![image](https://github.com/user-attachments/assets/2621028b-ad4c-4254-9934-459d32141a36)



### Demo
![Alt Text](https://github.com/limbo666/Shortcut_Dashboard/blob/main/images/ShortcutDashboard%20Demo.gif)

## Author

© Hand Water Pump - Nikos Georgousis

## License

Attribution License (CC BY)


