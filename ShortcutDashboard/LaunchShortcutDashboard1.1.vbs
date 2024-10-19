Option Explicit

On Error Resume Next

Dim objFSO, objShell, objFolder, objFile, strPath, strCommand, strLogFile
Dim strPattern, strVersion, highestVersion, highestVersionFile
Dim arrVersionParts, i
Dim LogEnabled

LogEnabled = 0 ' Enable log by setting this to 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
If Err.Number <> 0 Then
    WScript.Echo "Error creating FileSystemObject: " & Err.Description
    WScript.Quit
End If

Set objShell = CreateObject("WScript.Shell")
If Err.Number <> 0 Then
    WScript.Echo "Error creating WScript.Shell: " & Err.Description
    WScript.Quit
End If

strPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
strLogFile = strPath & "\launchlog.txt"
strPattern = "ShortcutDashboard(\d+\.\d+)\.ps1$"

' Delete existing log file if it exists
If objFSO.FileExists(strLogFile) Then
    objFSO.DeleteFile(strLogFile)
    If Err.Number <> 0 Then
        WScript.Echo "Error deleting existing log file: " & Err.Description
        Err.Clear
    End If
End If

' Function to write to log file
Sub WriteLog(strMessage)
   
   On Error Resume Next
if LogEnabled = 1 then
    Dim objLogFile
    Set objLogFile = objFSO.OpenTextFile(strLogFile, 8, True)
    If Err.Number <> 0 Then
        WScript.Echo "Error opening log file: " & Err.Description
        Err.Clear
    Else
        objLogFile.WriteLine Now & " - " & strMessage
        objLogFile.Close
    End If
End if
	
End Sub

WriteLog "Script started"
WriteLog "Scanning folder: " & strPath

highestVersion = "0.0"
Set objFolder = objFSO.GetFolder(strPath)
If Err.Number <> 0 Then
    WriteLog "Error getting folder: " & Err.Description
    WScript.Echo "Error getting folder: " & Err.Description
    WScript.Quit
End If

For Each objFile In objFolder.Files
    WriteLog "Checking file: " & objFile.Name
    If LCase(objFSO.GetExtensionName(objFile.Name)) = "ps1" Then
        WriteLog "File is a .ps1 file"
        If RegExpTest(objFile.Name, strPattern) Then
            strVersion = RegExpExecute(objFile.Name, strPattern, 0)
            WriteLog "File matches pattern. Version: " & strVersion
            
            If CompareVersions(strVersion, highestVersion) > 0 Then
                highestVersion = strVersion
                highestVersionFile = objFile.Name
                WriteLog "New highest version found: " & highestVersionFile
            End If
        Else
            WriteLog "File does not match pattern"
        End If
    Else
        WriteLog "File is not a .ps1 file"
    End If
Next

If highestVersionFile <> "" Then
    WriteLog "Highest version found: " & highestVersionFile
    strCommand = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & strPath & "\" & highestVersionFile & """"
    WriteLog "Launching command: " & strCommand
    objShell.Run strCommand, 0, False
    If Err.Number <> 0 Then
        WriteLog "Error launching PowerShell script: " & Err.Description
        WScript.Echo "Error launching PowerShell script: " & Err.Description
    Else
        WriteLog "PowerShell script launched successfully"
    End If
Else
    WriteLog "No matching files found."
    WScript.Echo "No ShortcutDashboard*.ps1 files found in the current directory."
End If

' Function to test if a string matches a regular expression
Function RegExpTest(strValue, strPattern)
    On Error Resume Next
    Dim objRegExp
    Set objRegExp = New RegExp
    objRegExp.Pattern = strPattern
    RegExpTest = objRegExp.Test(strValue)
    If Err.Number <> 0 Then
        WriteLog "Error in RegExpTest: " & Err.Description
        RegExpTest = False
    End If
End Function

' Function to execute a regular expression and return a matched group
Function RegExpExecute(strValue, strPattern, groupIndex)
    On Error Resume Next
    Dim objRegExp, matches
    Set objRegExp = New RegExp
    objRegExp.Pattern = strPattern
    Set matches = objRegExp.Execute(strValue)
    If Err.Number <> 0 Then
        WriteLog "Error in RegExpExecute: " & Err.Description
        RegExpExecute = ""
    ElseIf matches.Count > 0 Then
        RegExpExecute = matches(0).SubMatches(groupIndex)
    Else
        RegExpExecute = ""
    End If
End Function

' Function to compare two version strings
Function CompareVersions(version1, version2)
    On Error Resume Next
    Dim arr1, arr2, i
    arr1 = Split(version1, ".")
    arr2 = Split(version2, ".")
    
    For i = 0 To UBound(arr1)
        If CInt(arr1(i)) > CInt(arr2(i)) Then
            CompareVersions = 1
            Exit Function
        ElseIf CInt(arr1(i)) < CInt(arr2(i)) Then
            CompareVersions = -1
            Exit Function
        End If
    Next
    
    CompareVersions = 0
    If Err.Number <> 0 Then
        WriteLog "Error in CompareVersions: " & Err.Description
        CompareVersions = 0
    End If
End Function

WriteLog "Script finished"