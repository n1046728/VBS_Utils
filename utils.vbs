On error resume Next
'Logger setting--- 
Dim fs, f , myDate
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.OpenTextFile("log.txt", 8, -2)

'-----------------
'main
'-----------------
LogInfo("util.vbs start--------")
myFuncName = WScript.Arguments.Item(0)

If myFuncName = "GetDate" Then
	Wscript.Echo Eval(myFuncName)
End If

If myFuncName = "GetPreEOM" Then
	Wscript.Echo Eval(myFuncName)
End If

If myFuncName = "WebService" Then
	myFuncName = myFuncName & "(""" & WScript.Arguments.Item(1) & """,""" & WScript.Arguments.Item(2) & """)"
	Eval(myFuncName)
End If

If myFuncName = "SendEmail" Then
	If WScript.Arguments.Count = 1 Then
		myFuncName = myFuncName & "(""S"")"
	Else
		myFuncName = myFuncName & "(""F"")"
	End if
	Eval(myFuncName)
End If


'Error Handler
If err.Number <> 0 Then
	errNum = Err.Number
	errDesc = Err.Description
	LogErr("Call main()-ErrorHandler ,ErrorCode:" & errNum & " - ErroMsg:"& errDesc)
End If

LogInfo("util.vbs end----------")
f.Close


'-----------------
'Function
'-----------------

'return today yyyy/mm/dd
Function GetDate()

	rtnStr = Year(date) & "/"
	
	If Len(Month(date)) = 1 Then
		rtnStr = rtnStr & "0" & Month(date)  & "/"
	Else
		rtnStr = rtnStr & Month(date)  & "/"
	End If	
	
	If Len(Day(date)) = 1 Then
		rtnStr = rtnStr & "0" & Day(date)
	Else
		rtnStr = rtnStr & Day(date)
	End If	
	LogInfo("Call GetDate(),Return:" & rtnStr)
	GetDate = rtnStr
	
End Function

'return previou end of month day yyyymmdd
Function GetPreEOM()
	LastDay = DateSerial(Year(Date),Month(Date),0)
	myYear = Left(LastDay,4)
	myMonth = Mid(LastDay,InStr(LastDay,"/")+1,InStrRev(LastDay,"/") - InStr(LastDay,"/")-1)
	myDay = Right(LastDay, Len(LastDay)-InStrRev(LastDay,"/"))

	rtnStr = myYear & "/"
	If Len(myMonth) = 1 Then
		rtnStr = rtnStr & "0" & myMonth  & "/"
	Else
		rtnStr = rtnStr & myMonth  & "/"
	End If	
	
	If Len(myDay) = 1 Then
		rtnStr = rtnStr & "0" & myDay
	Else
		rtnStr = rtnStr & myDay
	End If	
	LogInfo("Call GetPreEOM(),Return:" & rtnStr)
	GetPreEOM = rtnStr
	
End Function

Function PortiaWebService(positionDate,batchType)
	Set fso = CreateObject ("Scripting.FileSystemObject")
	Set stdout = fso.GetStandardStream (1)
	Set stderr = fso.GetStandardStream (2)
	Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
	positionDate = Replace(positionDate,"'","")

	LogInfo("Call PortiaWebService(),positionDate:" & positionDate & ",batchType:" & batchType)
	
	URL = "http://127.0.0.1:8080/MyApp/rest/batch/"
	
	select case batchType
		case "D"
			URL = URL & "JobDaily"
		case "M"
			URL = URL & "JobMonthly"
		case else
			URL = ""	
	end select
	LogInfo("Call WebService(),URL:" & URL)
	
	objHTTP.Open "POST", URL, False
	objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0"
	objHTTP.setRequestHeader "Authorization", "Basic base64encodeduserandpassword"
	objHTTP.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
	objHTTP.setRequestHeader "CharSet", "charset=UTF-8"
	objHTTP.setRequestHeader "Accept", "application/json"

	' Send the json in correct format
	json = "{""positionDate"":""" & positionDate &"""}"
	LogInfo("Call PortiaWebService(),json:" & json)
	objHTTP.send (json)

	' Output error message to std-error and happy message to std-out. Should
	' simplify error checking
	If objHTTP.Status >= 400 And objHTTP.Status <= 599 Then
		'stderr.WriteLine "Error Occurred : " & objHTTP.status & " - " & objHTTP.statusText
		PortiaWebService = false
	Else
		'MsgBox objHTTP.Status
		'stdout.WriteLine "Success : " & objHTTP.status & " - " & objHTTP.ResponseText
		PortiaWebService = true
	End If
	
	LogInfo("Call PortiaWebService(),WS-Result:" & objHTTP.status & " - " & objHTTP.ResponseText)
	
End Function

Function SendEmail(sendType)
	LogInfo("Call SendEmail(),sendType:" & sendType)
	
	Set objConf = CreateObject("CDO.Configuration")
	Set objFlds = objConf.Fields
	objFlds.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPort
	objFlds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.chinalife.com.tw" 'your smtp server domain or IP address goes here
	objFlds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 'default port for email
	objFlds.Update
	
	'READ LOG INFO
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set file = fso.OpenTextFile("ERRORLOG_tail_20.log", 1)
	content = file.ReadAll
	
	'CREATE EMAIL 
	Set objEmail = CreateObject("CDO.Message")
	objEmail.Configuration = objConf
	objEmail.From = "JobReminder@chinalife.com.tw"
	objEmail.To = "test@gmail.com.tw"	
	fixedSubject = "Job Result"
	
	select case sendType
	case "S"
		objEmail.Subject = fixedSubject
		objEmail.Textbody = content
	case "F"
		objEmail.BodyPart.Charset = "utf-8"
		objEmail.Subject = fixedSubject & "_FAIL"
		objEmail.HTMLBody = "<h1 style=""color:red;"">ERROR! ERROR!</h1>" & Replace(content, vbCrlf, "<br>")
	end select
	
	objEmail.Send
	LogInfo("Call SendEmail(),EmailSubject:" & fixedSubject)
End Function

Private Function timestamp(d)
  ymd=fix(d)
  h=hour(d)
  m=minute(d)
  s=second(d)
  timestamp= year(d) &"/"& month(d) & "/" & day(d) &" "& right("00" & h,2)&":"& right("00" & m,2)& ":" & right("00" & s,2)& "."& _
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
