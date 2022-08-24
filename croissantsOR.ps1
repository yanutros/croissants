Start-Process "C:\Program Files\Rocket.Chat+\Rocket.Chat.exe"
Start-Sleep 3
$wshell
Start-Sleep 1
$wshell = New-Object -ComObject wscript.shell;
$wshell.AppActivate('Cosium.Chat')
Sleep 1
$wshell.SendKeys("^(k)")
Sleep 1
$wshell.SendKeys("fpa")
Sleep 1
$wshell.SendKeys("{ENTER}")
Sleep 1
$wshell.SendKeys(":croissant: i wan't des croissants")
Sleep 2
$wshell.SendKeys("{ENTER}")
Sleep 1
