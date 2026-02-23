$WshShell = New-Object -comObject WScript.Shell
$Desktop = $WshShell.SpecialFolders.Item("Desktop")
$Shortcut = $WshShell.CreateShortcut("$Desktop\BotEliva.lnk")
$Shortcut.TargetPath = "$PSScriptRoot\dist\IntelbrasWatcher.exe"
$Shortcut.WorkingDirectory = "$PSScriptRoot\dist"
$Shortcut.IconLocation = "$PSScriptRoot\dist\IntelbrasWatcher.exe,0"
$Shortcut.Save()
Write-Host "Atalho criado na Área de Trabalho com sucesso!"
