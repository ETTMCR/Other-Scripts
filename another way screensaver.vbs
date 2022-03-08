Option Explicit
Const ForAppending = 8
' writes log file, including absolute timestamp, after executing exe file

Dim  objFSO 
Dim WshShell
Dim objTextFile 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("D:\log.txt", ForAppending, True)
Set WshShell = WScript.CreateObject("WScript.Shell")


Dim dd, mm, yy, hh, nn, ss , datevalue, timevalue, dtsnow, dtsvalue ,milliseconds , t, temp , dot
dot = "." 

' MMS SECTION !!!!!!!!!!!!!!!!!
' PPT #1
  
	WshShell.AppActivate("MMS Arbel - v.6.4.0.30.36.2")' see the descending process in task manager 
	WshShell.SendKeys " " ' WORKS !!!!

' PPT #2	
	t = Timer
	temp = Int(t)
	milliseconds = Int((t-temp) * 1000)
	'Store DateTimeStamp once.
	dtsnow = Now()
	'Individual date components
	hh = Right("00" & Hour(dtsnow), 2)
	nn = Right("00" & Minute(dtsnow), 2)
	ss = Right("00" & Second(dtsnow), 2)
	'Build the time string in the format hh:mm:ss
	timevalue = hh & ":" & nn & ":" & ss
	'Concatenate both together to build the timestamp yyyy-mm-dd hh:mm:ss
	'dtsvalue = datevalue & " " & timevalue & " " &  dot & " " & milliseconds 
	dtsvalue = timevalue &  dot  & milliseconds  & " -MMS"
	objTextFile.WriteLine (dtsvalue)

'PSY SECTION !!!!!!!!!!!!!!!!!
' PPT #3
WshShell.AppActivate("Psy 4.68 (after Mar 17 2021) W7-VC++2010 CP=ROTEM-PC")
WshShell.SendKeys " "

' PPT #4	
	t = Timer
	temp = Int(t)
	milliseconds = Int((t-temp) * 1000)
	'Store DateTimeStamp once.
	dtsnow = Now()
	'Individual date components
	hh = Right("00" & Hour(dtsnow), 2)
	nn = Right("00" & Minute(dtsnow), 2)
	ss = Right("00" & Second(dtsnow), 2)
	'Build the time string in the format hh:mm:ss
	timevalue = hh & ":" & nn & ":" & ss
	'Concatenate both together to build the timestamp yyyy-mm-dd hh:mm:ss
	'dtsvalue = datevalue & " " & timevalue & " " &  dot & " " & milliseconds 
	dtsvalue = timevalue &  dot  & milliseconds  & " -PSY"
	objTextFile.WriteLine (dtsvalue)
	

objTextFile.Close 
WshShell.Run ("C:\Windows\System32\PhotoScreensaver.scr")  
