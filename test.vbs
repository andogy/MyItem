set WshShell = CreateObject("WScript.Shell")
WshShell.Run "%SystemRoot%\System32\SndVol.exe"
WScript.Sleep 1000
WshShell.AppActivate "Volume Mixer"
WshShell.SendKeys "{PGUP}" '
WshShell.SendKeys "{PGUP}" '
WshShell.SendKeys "{PGUP}" '
WshShell.SendKeys "{PGUP}" '
WshShell.SendKeys "{PGUP}" '
WshShell.SendKeys "%{F4}"  '