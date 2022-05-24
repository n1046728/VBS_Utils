On error resume Next
'Logger setting---
Dim fs, f 
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.OpenTextFile("log.txt", 8, -2)
'-----------------

LogInfo("start calculate")
WScript.Sleep 100
a = 1/0


If err.Number <> 0 Then
errNum = Err.Number
errDesc = Err.Description

LogErr(errNum & " - "& errDesc)


end if

f.Close




Private Function timestamp(d)
  ymd=fix(d)
  h=hour(d)
  m=minute(d)
  s=second(d)
  timestamp=year(d) &"/"& month(d) & "/" & day(d) &" "& right("00" & h,2)&":"& right("00" & m,2)& ":" & right("00" & s,2)& "."& _
  right("000"& 86400000*(d-timeserial(h,m,s)-ymd),3)
End Function  

private Function LogInfo(msg)
	d1 = date + timer/86400
	f.writeLine timestamp(d1) & " ||INFO || " & msg
End function

private Function LogErr(msg)
	d1 = date + timer/86400
	f.writeLine timestamp(d1) & " ||ERROR|| " & msg
End function