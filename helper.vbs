' Copyright 2014 Towry Wang <http://towry.me>

' read file content
Private Function ReadFile(filename)
	Set fs = CreateObject("Scripting.FileSystemObject")

	Dim currentPath, fullPath
	Dim sWords, sSearch, sContent

	currentPath = fs.GetFolder(".").Path
	fullPath = currentPath & "/" & filename

	If fs.FileExists(fullPath) Then
		 Set ts = fs.OpenTextFile(fullPath, 1, false)
		 sContent = ts.ReadAll()
		 ts.Close()
	Else
		msgbox("File not exists")
		WScript.Quit(1)
	End If

	Set fs = Nothing
	ReadFile = sContent
End Function

Private Function AbsolutePath(filename)
	Set fs = CreateObject("Scripting.FileSystemObject")

	Dim currentPath, fullPath

	currentPath = fs.GetFolder(".").Path
	fullPath = currentPath & "/" & filename

	Set fs = Nothing
	AbsolutePath = fullPath
End Function

' read from registry
Private Function readFromRegistry(strRegistryKey, strDefault)
	Dim ws, value

	On Error Resume Next
	Set ws = CreateObject("WScript.Shell")
	value = ws.RegRead(strRegistryKey)

	If Err.number <> 0 Then
		WScript.Echo Err.description
		WScript.Quit(1)
	Else
		readFromRegistry = value
	End If

	Set ws = Nothing
End Function

' write to registry
Private Function writeToRegistry(strRegistryKey, value)
	Dim ws, bkey

	On Error Resume Next
	Set ws = CreateObject("WScript.Shell")
	ws.RegWrite strRegistryKey, value, "REG_SZ"

	If Err.number <> 0 Then
		WScript.Echo Err.description
		WScript.Quit(1)
	End If
End Function

' get key,value pair from a line 
Private Function Pair(strLine)
	Dim key, value, strChar, bEqualFlag, bWSFlag
	bEqualFlag = False 
	bWSFlag = False

	For i=1 To Len(strLine)
		strChar = Mid(strLine, i, 1)
		If bEqualFlag <> True And Trim(strChar) <> "" And strChar <> Chr(9) Then
			If strChar = "=" Then
				bEqualFlag = True
			Else
				key = key & strChar
			End IF
		ElseIf bEqualFlag = True Then
			If Trim(strChar) = "" And bWSFlag <> True Then
				bWSFlag = True
			Elseif bWSFlag = True Then
				value = value & strChar
			End If 
		End If 
	Next

	' we need a {=} in that string
	If bEqualFlag <> True Then
		MsgBox("Error 91#\n File format error.")
		WScript.Quit(1)
	End If

	key = Trim(key)
	value = Trim(value)

	Pair = array(key, value)
End Function

' Write to path 
Private Function writeToPath(strPathValue)
	Set ws = CreateObject("WScript.Shell")
	Set env = ws.Environment("System")
	Dim path 

	path = env("PATH")
	path = path & ";" & strPathValue
	env("PATH") = path

End Function

Private Function removeFromPath(strPathValue)
	Set ws = CreateObject("WScript.Shell")
	Set env = ws.Environment("System")
	Dim path

	path = env("PATH")
	path = Replace(path, ";" & strPathValue, "")
	env("PATH") = path
End Function

' main entry
Public Function Main(filename)
	' str = readFromRegistry("HKEY_CLASSES_ROOT\http\shell\open\command\", "Nothing")
	' WScript.Echo "returned " & str
	' keyname = "HKEY_CURRENT_USER\SOFTWARE\Classes\http\shell\open\command\"
	' value = "'D:\Program Files\Opera\launcher.exe' -noautoupdate -- '%1'"

	' writeToRegistry keyname, value

	Dim abs, nJobCount, sJobState

	nJobCount = 0
	abs = AbsolutePath(filename)
	Set fs = CreateObject("Scripting.FileSystemObject")

	If Not fs.FileExists(abs) Then
		msgbox("File not exists")
		WScript.Quit(1)
	End If

	Set ts = fs.OpenTextFile(abs, 1, false)

	Do Until ts.AtEndOfStream
		Dim strLine, strStart, section
		Dim Gname, Gvalue, Gtarget 

		strLine = ts.ReadLine
		strStart = Mid(strLine, 1, 1)

		' #check get section
		If strStart = "[" Then
			sJobState = "pending"
			For i=2 To Len(strLine)
				strChar = Mid(strLine, i, 1)
				If strChar = "]" Then
					Exit For
				End If
				section = section & strChar
			Next
		End If

		If strStart <> "[" Then
			' #check If section is nothing, than there must be something wrong
			If sJobState = "pending" And section = "" Then
				MsgBox("Error 1# File format error.")
				WScript.Quit(1)
			ElseIf sJobState = "completed" Then
				sJobState = "started"
			End If

			' else it's a <key,value> pair
			' below depend on the section value
			' if section is a registry
			If LCase(section) <> "" Then
				' we need key#opt, value, target
				Dim aPair, key, value, target

				aPair = Pair(strLine)
				key = LCase(aPair(0))
				value = aPair(1)

				If key <> "" And key = "name" Then 
					Gname = value
				ElseIf key <> "" And key = "value" Then
					Gvalue = value
				ElseIf key <> "" And key = "target" Then
					Gtarget = value
				End If
			End IF
		End If

		If Gvalue <> "" And Gtarget <> "" And LCase(section) = "taskbar" Then
			retVal = pinToTaskBar(Gtarget, Gvalue)
			section = ""
			Gname = ""
			Gvalue = ""
			Gtarget = ""

			nJobCount = nJobCount + 1
			If retVal = "fail" Then
				nJobCount = nJobCount - 1
			End If
			sJobState = "completed"
		End If

		If Gvalue <> "" And LCase(section) = "path" Then
			' Only vlaue is required.
			writeToPath Gvalue
			section = ""
			Gname = ""
			Gvalue = ""
			Gtarget = ""

			nJobCount = nJobCount + 1
			sJobState = "completed"
		End If

		If Gvalue <> "" And Gtarget <> "" And LCase(section) = "registry" Then
			' finish last job
			' do the registry job
			' value and target is required.
			writeToRegistry Gtarget & Gname, Gvalue
			section = ""
			Gname = ""
			Gvalue = ""
			Gtarget = ""
			
			nJobCount = nJobCount + 1
			sJobState = "completed"
			' Done
		End If
	Loop

	ts.Close
	Set ts = Nothing
	Set fs = Nothing

	WScript.Echo CStr(nJobCount) & " jobs done."
End Function

' run
' maybe add a feature to select config file?
Main "setting.ini"
