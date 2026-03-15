$WshShell = New-Object -comObject WScript.Shell
$startupPath = [System.Environment]::GetFolderPath('Startup')
$lnk = $WshShell.CreateShortcut($startupPath + '\CNC_자동실행.lnk')
$lnk.TargetPath = 'C:\Users\admin\Desktop\work\자동화\CNC분석\자동실행.bat'
$lnk.WorkingDirectory = 'C:\Users\admin\Desktop\work\자동화\CNC분석'
$lnk.WindowStyle = 1
$lnk.Save()
Write-Host "등록 완료: $startupPath\CNC_자동실행.lnk"
