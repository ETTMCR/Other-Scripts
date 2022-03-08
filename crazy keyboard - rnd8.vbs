Dim max,min
MsgBox "Look at your keyboard for a cool light show!" 
MsgBox "to shut off ctrl+alt+del --> wscript.shell OR microsoft based..." 
set wshShell = wscript.CreateObject("wscript.shell")
wscript.Createobject("WScript.Shell")
Set objShell = CreateObject("Shell.Application")
objShell.ToggleDesktop

max=7
min=0

do
Randomize

wscript.sleep 150
myName = (Int((max-min+1)*Rnd+min))
Select Case myName
Case 0
Case 1
wshShell.sendkeys"{NUMLOCK}"

Case 2

wshShell.sendkeys"{CAPSLOCK}"

Case 3
wshShell.sendkeys"{NUMLOCK}"
wshShell.sendkeys"{CAPSLOCK}"

Case 4

wshShell.sendkeys"{SCROLLLOCK}"
Case 5
wshShell.sendkeys"{NUMLOCK}"

wshShell.sendkeys"{SCROLLLOCK}"
Case 6

wshShell.sendkeys"{CAPSLOCK}"
wshShell.sendkeys"{SCROLLLOCK}"
Case 7
wshShell.sendkeys"{NUMLOCK}"
wshShell.sendkeys"{CAPSLOCK}"
wshShell.sendkeys"{SCROLLLOCK}"
End Select


loop