Set WshShell= WScript.CreateObject("WScript.shell")
WshShell.AppActivate"±ø÷Ì"
for i=1 to 10
WScript.sleep 10
WshShell.SendKeys "^v"
WshShell.SendKeys i
WshShell.SendKeys "%s"
Next