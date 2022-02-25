$WshShell = New-Object -comObject WScript.Shell
$destination = $([Environment]::GetFolderPath("Desktop"))
$DestinationPath = Join-Path -Path $destination -ChildPath "\\Somatus Updater.lnk"
$SourceExe = "cmd.exe"
$Shortcut = $WshShell.CreateShortcut($DestinationPath)
$Shortcut.TargetPath = $SourceExe
$Shortcut.arguments = "/c ""C:\Windows\Temp\Somatus_Contact_Info_Updater.bat"""
$Shortcut.iconlocation = "https://iconarchive.com/download/i24563/mattahan/umicons/Letter-S.ico"
$Shortcut.Save()
