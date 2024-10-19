# INI file path
$iniFile = Join-Path $PSScriptRoot "options.ini"

# Function to read a value from the INI file
function Get-IniValue {
    param (
        [string]$section,
        [string]$key
    )
    $ini = Get-Content $iniFile -ErrorAction SilentlyContinue
    if ($ini) {
        $currentSection = ""
        foreach ($line in $ini) {
            if ($line -match '^\[(.+)\]$') {
                $currentSection = $matches[1]
            }
            elseif (($currentSection -eq $section) -and ($line -match "^$key\s*=\s*(.+)$")) {
                return $matches[1]
            }
        }
    }
    return $null
}



# File path of the log file
$logFile = "ShortcutDashboard_Log.txt"
# Get the LogEnabled value and convert it to an integer
$logEnabledValue = Get-IniValue -section "Settings" -key "LogEnabled"
$logEnabled = if ([int]::TryParse($logEnabledValue, [ref]$null)) { [int]$logEnabledValue } else { 0 }

# Add this line for debugging (you can remove it later)
Write-Host "LogEnabled value: $logEnabled"

# Log function for general log file
# Log function for general log file
function Write-Log {   
    param([string]$message)
    
    # Debug output inside the function
    Write-Host "Inside Write-Log: logEnabled = $logEnabled"
    Write-Host "Condition result: $($logEnabled -eq 1)"
    
    if ($logEnabled -eq 1) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $logFile -Value "$timestamp $message"
        Write-Host "Log written: $timestamp $message"
    } else {
        Write-Host "Log not written due to logEnabled being $logEnabled"
    }
}
# Logging the start of the script
Write-Log "SCRIPT: Script execution started."

# Define form size and properties
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$form = New-Object System.Windows.Forms.Form
# Get screen height dynamically
$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height

# Set maximum height allowed for the form (Screen height minus 80)
$maxFormHeight = $screenHeight - 80

# Create a Panel to hold dynamically created buttons
$panel = New-Object Windows.Forms.Panel
$panel.Dock = 'Fill'
$panel.AutoScroll = $true  # Enable automatic scroll when needed
$panel.BorderStyle = 'None'

$form.Text = "ShortcutDashboard"
$form.Width = 280  # Default form width, dynamically adjusted later
$form.Height = 100
$form.StartPosition = 'Manual'
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.MaximizeBox = $false


# Function to write a value to the INI file
function Set-IniValue {
    param (
        [string]$section,
        [string]$key,
        [string]$value
    )
    if (-not (Test-Path $iniFile)) {
        New-Item -ItemType File -Path $iniFile | Out-Null
        Add-Content -Path $iniFile -Value "[$section]"
        Add-Content -Path $iniFile -Value "$key = $value"
    }
    else {
        $content = Get-Content $iniFile
        $sectionFound = $false
        $keyFound = $false
        $newContent = @()
        foreach ($line in $content) {
            if ($line -match '^\[(.+)\]$') {
                if ($matches[1] -eq $section) {
                    $sectionFound = $true
                }
                elseif ($sectionFound -and -not $keyFound) {
                    $newContent += "$key = $value"
                    $keyFound = $true
                }
            }
            elseif ($sectionFound -and $line -match "^$key\s*=") {
                $line = "$key = $value"
                $keyFound = $true
            }
            $newContent += $line
        }
        if (-not $sectionFound) {
            $newContent += "[$section]"
            $newContent += "$key = $value"
        }
        elseif (-not $keyFound) {
            $newContent += "$key = $value"
        }
        Set-Content -Path $iniFile -Value $newContent
    }
    Write-Log "SCRIPT: Updated INI file: [$section] $key = $value"
}

# Function to save form position
function Save-FormPosition {
    $formLocation = "$($form.Left),$($form.Top)"
    Set-IniValue -section "Settings" -key "FormLocation" -value $formLocation
    Write-Log "SCRIPT: Form position saved to INI file: $formLocation"
}

# Function to load form position
function Load-FormPosition {
    $formLocation = Get-IniValue -section "Settings" -key "FormLocation"
    if ($formLocation) {
        $left, $top = $formLocation.Split(',')
        if ([int]$left -lt -10) {
		$left=0
		}
		   if ([int]$top -lt -10) {
		$top=0
		}
		$form.Left = [int]$left
        $form.Top = [int]$top
        Write-Log "SCRIPT: Form position loaded from INI file: $formLocation"
    }
    else {
        $form.StartPosition = 'CenterScreen'
        Write-Log "SCRIPT: No saved position found, using center screen"
    }
}

# Load form position
Load-FormPosition

# Set form icon
$iconPath = Join-Path $PSScriptRoot "icon.ico"
if (Test-Path $iconPath) {
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
    Write-Log "SCRIPT: Custom icon loaded from 'icon.ico'"
} else {
    $shell32 = Join-Path ([System.Environment]::SystemDirectory) "imageres.dll"
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($shell32)
    Write-Log "SCRIPT: Default system icon set from imageres.dll"
}

# ToolTip initialization
$toolTip = New-Object System.Windows.Forms.ToolTip

# Detect root path (location of this script)
$rootPath = $PSScriptRoot
$scriptsFolder = Join-Path $rootPath "Scripts"

# List of extensions to search for
$extensions = @("*.exe", "*.bat", "*.vbs", "*.cmd", "*.ps1", "*.lnk")

# Define a dictionary for file-specific icons
$fileIcons = @{
    ".exe" = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\shell32.dll")
    ".bat" = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\cmd.exe")
    ".vbs" = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\wscript.exe")
    ".cmd" = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\cmd.exe")
    ".ps1" = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe")
}

# Track the button position (initial position) as a script-scoped variable
$script:buttonTop = 10
$buttonHeight = 30
$buttonPadding = 5
$maxButtonWidth = 240  # Set a minimum button width

# Function to execute the file when a button is clicked
function Execute-File {
    param($filePath, $arguments)

    try {
        # Log user action
        Write-Log "USER: Attempting to execute file '$filePath' with arguments '$arguments'"

        # Determine the file extension and execute accordingly
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()

        if ($extension -eq ".ps1") {
            # Execute PowerShell scripts with powershell.exe
            Write-Log "SCRIPT: Executing PowerShell script: $filePath $arguments"
          #  Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$filePath`" $arguments"  #-NoNewWindow
		 # Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-ExecutionPolicy Bypass", "-File `"$filePath`"", $arguments
			Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit -ExecutionPolicy Bypass -File `"$filePath`" $arguments"  #-NoNewWindow
        } elseif ($extension -eq ".lnk") {
            # For shortcuts, resolve the target and execute it with arguments
            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($filePath)
            $targetPath = $shortcut.TargetPath
            $shortcutArguments = $shortcut.Arguments
            $fullArguments = "$shortcutArguments $arguments".Trim()
            Write-Log "SCRIPT: Executing shortcut target: $targetPath with arguments: $fullArguments"
            Start-Process -FilePath $targetPath -ArgumentList $fullArguments
        } else {
            # For all other file types, use the default system association with arguments
            Write-Log "SCRIPT: Executing file with default association: $filePath $arguments"
            Start-Process -FilePath $filePath -ArgumentList $arguments
        }
    } catch {
        Write-Log "SCRIPT: Error executing '$filePath' with arguments '$arguments' - $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error: Unable to execute the file '$filePath'.", "Execution Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Function to add a button dynamically
function Add-Button {
    param($file)
    
    $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($file)
    $fileExtension = [System.IO.Path]::GetExtension($file).ToLower()

    # Create the button
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $fileNameWithoutExtension
    $button.Height = $buttonHeight
    $button.Top = $script:buttonTop
    $button.Left = 10  # Padding for form margin
    $button.Width = $maxButtonWidth  # Set a consistent button width for all

    # Set tooltip to show the full file path
    $toolTip.SetToolTip($button, $file)

    # Add icon if available for the file type
    if ($fileExtension -eq ".exe" -or $fileExtension -eq ".lnk") {
        # Use the specific executable's icon or shortcut's icon
        $button.Image = [System.Drawing.Icon]::ExtractAssociatedIcon($file).ToBitmap()
        Write-Log "SCRIPT: Loaded icon for file: $file"
    } elseif ($fileIcons.ContainsKey($fileExtension)) {
        $button.Image = $fileIcons[$fileExtension].ToBitmap()
    }

    # Ensure proper alignment of text and icon
    $button.TextAlign = "MiddleRight"
    $button.ImageAlign = "MiddleLeft"

    # Set the action when clicked, passing the captured file path directly
    $button.Add_Click({
        $clickedFile = $this.Tag
        Write-Log "USER: Button clicked for file '$clickedFile'"
        if ($fileExtension -eq ".lnk") {
            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($clickedFile)
            Execute-File $shortcut.TargetPath $shortcut.Arguments
        } else {
            Execute-File $clickedFile
        }
    })
    $button.Tag = $file

    # Add the button to the form
    $panel.Controls.Add($button)

    # Move the position for the next button
    $script:buttonTop += $buttonHeight + $buttonPadding
}
$form.Controls.Add($panel)

# Dynamically adjust form height if it exceeds screen height - 80
if ($form.Height -gt $maxFormHeight) {
    $form.Height = $maxFormHeight
}
# Scan for files in the Scripts folder
try {
    Write-Log "SCRIPT: Scanning folder '$scriptsFolder' for files."

    # Check if Scripts folder exists, if not, create it
    if (-not (Test-Path $scriptsFolder)) {
        Write-Log "SCRIPT: 'Scripts' folder missing. Creating the folder."
        New-Item -Path $scriptsFolder -ItemType Directory | Out-Null
        Write-Log "SCRIPT: 'Scripts' folder created."
    }

    $files = Get-ChildItem -Path $scriptsFolder -Recurse -Include $extensions

    if ($files.Count -eq 0) {
        Write-Log "SCRIPT: No files found in '$scriptsFolder'."
        [System.Windows.Forms.MessageBox]::Show("No executable scripts found in the 'Scripts' folder.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        foreach ($file in $files) {
            Write-Log "SCRIPT: Detected file '$($file.FullName)'"
            # Skip this script itself
            if ($file.Name -ne "ShortcutDashboard.ps1") {
                Add-Button -file $file.FullName
            }
        }

        # Add copyright label
        $copyrightLabel = New-Object System.Windows.Forms.Label
        $copyrightLabel.Text = "© Hand Water Pump 2025"
        $copyrightLabel.AutoSize = $true
        $copyrightLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $copyrightLabel.Width = $form.ClientSize.Width
        $copyrightLabel.Left = 0
        $copyrightLabel.Top = $script:buttonTop

        $panel.Controls.Add($copyrightLabel)

        # Adjust form height dynamically to fit all buttons and copyright label
        $formHeight = $script:buttonTop + $copyrightLabel.Height + 40
        $screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height
        $form.Height = [Math]::Min($formHeight, $screenHeight)
        Write-Log "SCRIPT: Form height set to $($form.Height)"

        # Center the copyright label
        $copyrightLabel.Left = ($form.ClientSize.Width - $copyrightLabel.Width) / 2
    }
} catch {
    Write-Log "SCRIPT: Error while scanning folder - $($_.Exception.Message)"
    [System.Windows.Forms.MessageBox]::Show("An error occurred while scanning the folder.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}

# Save form position when closing
$form.Add_FormClosing({ Save-FormPosition })

# Display the form
try {
    Write-Log "SCRIPT: Launching the form."
    $form.ShowDialog() | Out-Null
} catch {
    Write-Log "SCRIPT: Error displaying the form - $($_.Exception.Message)"
    [System.Windows.Forms.MessageBox]::Show("An error occurred while displaying the form.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}

# Logging script completion
Write-Log "SCRIPT: Script execution completed."