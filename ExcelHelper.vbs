'Exit code :
'0 : Operation Success.
'1 : Unknown error occured
'10: Excel File Not Opened
'11: Excel File Is Opened
'12: Number of Arguments doesn't match
'13: File Not Exists
'14" File Exists

Public Sub IsExcelOpen(ExcelFileName)
	On Error Resume Next
	'First check if a file exists
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	if not fso.FileExists(ExcelFileName) then
		fso = Nothing
		Wscript.Quit 13
	end if

	Set objExcel = GetObject(, "Excel.Application")  'attach to running Excel instance
	If Err Then
	  If Err.Number = 429 Then
	  	'"Workbook not open (Excel is not running)."
	    WScript.Quit 10
	  End If
	End If
	On Error Goto 0

	Set wb = Nothing
	For Each obj In objExcel.Workbooks
	  If obj.Name = ExcelFileName Then  'use obj.FullName for full path
	    Set wb = obj
	    WScript.Quit 11
	    Exit For
	  End If
	Next

	If wb Is Nothing Then
		'ExcelFileName is not opened
		WScript.Quit 10
	End If
End Sub

Public Sub CheckFileExist(FileName)
	dim fso
	set fso = CreateObject("Scripting.FileSystemObject")
	if fso.FileExists(FileName) then
		'fso = Nothing
		Wscript.Quit 14
	else
		'fso = Nothing
		Wscript.Quit 13 
	End If
end sub

Public Sub CheckExcelOpen(ExcelFileName)

	On Error Resume Next
	Set objExcel = GetObject(, "Excel.Application")  'attach to running Excel instance
	If Err Then
	  If Err.Number = 429 Then
	  	'"Workbook not open (Excel is not running)."
	    WScript.Quit 10
	  End If
	End If
	On Error Goto 0

	Set wb = Nothing
	For Each obj In objExcel.Workbooks
	  If obj.Name = "ExcelHelper.xls" Then  'use obj.FullName for full path
	    Set wb = obj
	    WScript.Quit 11
	    Exit For
	  End If
	Next

	If wb Is Nothing Then
		'ExcelFileName is not opened
		WScript.Quit 10
	End If	
End Sub


Public Sub CloseWorkBook(strName)

	Dim i
	Dim XLAppFx
	Dim NotOpen
	Dim TargetWorkbook
     
     'Find/create an Excel instance
    On Error Resume Next
    Set XLAppFx = GetObject(, "Excel.Application")
    If Err.Number = 429 Then
        NotOpen = True
        Set XLAppFx = CreateObject("Excel.Application")
        Err.Clear
    End If
     
     'Loop through all open workbooks in such instance
    For i = XLAppFx.Workbooks.Count To 1 Step -1
        If XLAppFx.Workbooks(i).Name = strName Then 
			Set TargetWorkbook = XLAppFx.Workbooks(i)
			Exit For
		End If
    Next
	
	If(Not IsNull(TargetWorkbook)) Then
		TargetWorkbook.Close(False)
	End If
	
End Sub



Public Sub CloseExcel(ExcelFileName)
	On Error Resume Next
	Set objExcel = GetObject(, "Excel.Application")
	If Err Then
	  If Err.Number = 429 Then
	    WScript.Echo "Workbook not open (Excel is not running)."
	  Else
	    WScript.Echo Err.Description & " (0x" & Hex(Err.Number) & ")"
	    WScript.Quit 1
	  End If
	  WScript.Quit 1
	End If
	On Error Goto 0


	For Each obj In objExcel.Workbooks
	  If obj.Name = ExcelFileName Then  'use obj.FullName for full path
	  	obj.Save
	    if objExcel.Workbooks.Count = 1 then 
			objExcel.Quit
	    else 
	    	obj.Close
	    end if  
	    If Err Then 
	    	WScript.Echo Err.Description &  ExcelFileName & " is closed."
	    End IF
	    Exit For
	  End If
	Next
	Set objExcel = Nothing
	WScript.Quit 0	
End Sub


Public Sub delay (time)
	WScript.Sleep time*1000
	WScript.Quit 0
End Sub

Public Sub EmailSender (message)
	Set objMail = CreateObject("CDO.Message")
	Set objConf = CreateObject("CDO.Configuration")
	Set objFlds = objConf.Fields
	objFlds.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPort
	objFlds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com" 'your smtp server domain or IP address goes here
	objFlds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 'default port for email
	objFlds.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
	'uncomment next three lines if you need to use SMTP Authorization
	objFlds.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "lxp1991"
	objFlds.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "xxxxxxxx"
	objFlds.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'cdoBasic
	objFlds.Update
	objMail.Configuration = objConf
	objMail.From = "lxp1991@gmail.com"
	objMail.To = "lxp1991@gmail.com"
	objMail.Subject = message
	objMail.TextBody = message
	objMail.Send
	Set objFlds = Nothing
	Set objConf = Nothing
	Set objMail = Nothing
	WScript.Quit 0
End Sub

Public Sub Main()
	Dim arg
	Dim FileName
	Dim message
	Set args = Wscript.Arguments

	if args.Count <> 2 Then
		WScript.Quit 12
	End If

	if args.Item(0) = "IsExcelOpen" Then
		ExcelFileName = args.Item(1)
		IsExcelOpen(FileName)
	End IF

	if args.Item(0) = "delay" Then
		delay(args.Item(1))
	End If

	if args.Item(0) = "CloseExcel" Then
		FileName = args.Item(1)
		CloseExcel(args.Item(1))
	End If

	if args.Item(0) = "EmailSender" Then
		message = args.Item(1)
		EmailSender(message)
	end If

	if args.Item(0) = "CheckFileExist" Then 
		FileName = args.Item(1)
		CheckFileExist(FileName)
	End If
	WScript.Quit 0
End Sub


Main()