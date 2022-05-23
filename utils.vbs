
myFuncName = WScript.Arguments.Item(0)

If myFuncName = "GetDate" Then
	MsgBox "GetDate"
	Wscript.Echo Eval(myFuncName)
End If

If myFuncName = "PortiaWebService" Then
	MsgBox "PortiaWebService"
	myFuncName = myFuncName & "(" & WScript.Arguments.Item(1) & "," & WScript.Arguments.Item(2) & ")"
	MsgBox myFuncName
	Eval(myFuncName)
End If


'return yyyy/mm/dd
Function GetDate()

	rtnStr = Year(date) & "/"
	
	If Len(Month(date)) = 1 Then
		rtnStr = rtnStr & "0" & Month(date)  & "/"
	Else
		rtnStr = rtnStr & Month(date)  & "/"
	End If	
	
	If Len(Day(date)) = 1 Then
		rtnStr = rtnStr & "0" & Month(date)  & "/"
	Else
		rtnStr = rtnStr & Day(date)
	End If	
	'MsgBox rtnStr
	GetDate = rtnStr
	
End Function



Function PortiaWebService(positionDate,urlPosfix)

MsgBox positionDate + urlPosfix
End Function