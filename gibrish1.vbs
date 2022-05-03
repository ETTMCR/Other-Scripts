
Set WshShell = WScript.CreateObject ("Wscript.Shell")
WshShell.Run "www.google.com"
Dim copied
copied = CreateObject("htmlfile").ParentWindow.ClipboardData.Getdata("text")
Wscript.Sleep 500 ' option to be 750
WshShell.sendkeys copied ' & copied
Wscript.Sleep 500 ' option to be 750
WshShell.sendkeys "{Enter}" 'stroke to google search
Wscript.Sleep 250
For Counter = 1 To 19
	Wscript.Sleep 50
	WshShell.sendkeys "{TAB}"
Next

WshShell.sendkeys "{Enter}" ' stroke the new suggestion
Wscript.Sleep 250

 
For Counter = 1 To 14 ' go back to search inputbox
	Wscript.Sleep 50
	WshShell.sendkeys "+{TAB}" 'shift+tab
Next
Wscript.Sleep 50
WshShell.SendKeys "^c" ' copy the new suggestion
Wscript.Sleep 50

WshShell.sendkeys "%{TAB}" ' retruns to last program by alt+tab
Wscript.Sleep 100
WshShell.SendKeys "^v" ' paste the new suggestion

Wscript.Quit
'akuo gkhfo
'= ,שלום עליכם
'akuo gkhfo nktfh gkhui
'=שלום עליכם מלאכי עליון
