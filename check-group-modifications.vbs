''
''	
''	Check Group Modifications
''
''	Scheduled to run every 1 hour
'' 	1) Rename the current.txt file to last.txt
''	2) Write the current members to a file current.txt
''	3) If the size of current.txt <> last.txt then there was a change in the group members
''	4) Send e-mail to the security officer
''
''
''		Function EncloseWithDQ
''		Sub CheckGroupForChanges
''		Function FileExists
''		Function GetFileSize
''		Sub DeleteFile
''		Sub RenameFile


Option Explicit




Dim	gstrGroupDn





Sub DeleteFile(sPath)
	''
	''	DeleteFile()
	''	
	''	Delete a file specified as "d:\folder\filename.ext"
	''
	''	sPath	The name of the file to delete.
	''
   	Dim oFSO
   	
   	Set oFSO = CreateObject("Scripting.FileSystemObject")
   	If oFSO.FileExists(sPath) Then
   		oFSO.DeleteFile sPath, True
   	End If
   	Set oFSO = Nothing
End Sub '' DeleteFile


Function EncloseWithDQ(ByVal s)
	''
	''	Returns an enclosed string s with double quotes around it.
	''	Check for exising quotes before adding adding.
	''
	''	s > "s"
	''
	
	If Left(s, 1) <> Chr(34) Then
		s = Chr(34) & s
	End If
	
	If Right(s, 1) <> Chr(34) Then
		s = s & Chr(34)
	End If

	EncloseWithDQ = s
End Function '' of Function EncloseWithDQ




Function FileExists(sPath)
	'==
	'==	Check for the existence of a file.
	'==
	'==	Variables:
	'==		sPath	The full path of the file to check
	'==
	'==	Returns:
	'==		True	file exists
	'==		False	The file doesn't exist.
	'==
	
	Dim	oFso
	Dim	bReturn
	
	Set oFso = CreateObject("Scripting.FileSystemObject")
	
	If oFso.FileExists(sPath) = True Then
		bReturn = True
	Else
		bReturn = False
	End If
	
	FileExists = bReturn
	
	Set oFso = Nothing
End Function '== FileExists




Function GetFileSize(ByVal sFName)
	'
	'	GetFileSize
	'
	'	Return the length of a file or folder. Returns -1 when file or folder does not exist
	'

	Dim		objFso
	Dim		objFile
	
	GetFileSize = -1
	
	Set objFso = CreateObject("Scripting.FileSystemObject")
	
	If objFso.FileExists(sFName) = True Then
		Set objFile = objFso.GetFile(sFName)
		GetFileSize = objFile.Size
	End If
	
	Set objFile = Nothing
	Set objFso = Nothing
End Function


Sub RenameFile(ByVal strPathOrg, ByVal strPathNew)
	Dim		objFso
	
	If FileExists(strPathOrg) =  True Then
		'' The Org file exists, continue.
		If FileExists(strPathNew) = False Then
			'' The new file doesn't exists, continue.
			Set objFso = CreateObject("Scripting.FileSystemObject")
			Call objFso.MoveFile(strPathOrg, strPathNew)
			Set objFso = Nothing
		Else
			WScript.Echo "WARNING: File " & strPathNew & " already exists"
		End If
	End If
End Sub



Sub SendMail(ByVal strGroupCn, ByVal strPathCurr, ByVal strPathPrev)
	Dim		strMailTo
	Dim 	c
	Dim		objShell
	Dim		strBody
	
	strMailTo = "perry.vandenhondel@ns.nl"
	strBody = "body.txt"
	
	' blat readme.md -to perry.vandenhondel@ns.nl -f perry.vandenhondel@ns.nl -subject "TEST 1502"  -server vm70as005.rec.nsint -port 25

	c = "r:\tools\blat.exe " & strBody & " "
	c = c & "-to " & strMailTo & " "
	c = c & "-f " & "nsg.hostingadbeheer@ns.nl "
	c = c & "-subject " & EncloseWithDQ("CHANGE IN GROUP " & strGroupCn & " DETECTED") & " "
	c = c & "-attacht " & strPathCurr & " "
	c = c & "-attacht " & strPathPrev & " "
	c = c & "-server vm70as005.rec.nsint "
	c = c & "-port 25"
			
	Set objShell = CreateObject("WScript.Shell")
	WScript.Echo "RUNNING: " & c
	objShell.Run "cmd /c " & c, 0, True
	Set objShell = Nothing
End Sub



Sub CheckGroupForChanges(ByVal strRootDse, ByVal strGroupCn)
	Dim		c
	Dim		strPathCurr		'' Current export of group members
	Dim		strPathPrev		'' Previous export of group members
	Dim		objShell
	Dim		intSizeCurr
	Dim		intSizePrev
	
	WScript.Echo "Check group: " & strGroupCn

	strPathCurr = strGroupCn & "-curr.txt"
	strPathPrev = strGroupCn & "-prev.txt"
	
	WScript.Echo strPathCurr
	WScript.Echo strPathPrev
	
	'' file current.txt exist, rename to last.txt
	If FileExists(strPathPrev) =  True Then
		Call DeleteFile(strPathPrev)
	End If
	
	'' Rename the current.txt to last.txt
	Call RenameFile(strPathCurr, strPathPrev)
	
	
	'' 
	c = "adfind.exe -b " & EncloseWithDQ(strRootDse) & " -f " & EncloseWithDQ("CN=" & strGroupCn) & " member -list >" & strPathCurr
	Set objShell = CreateObject("WScript.Shell")
	WScript.Echo "RUNNING: " & c
	objShell.Run "cmd /c " & c, 0, True
	Set objShell = Nothing
	
	intSizeCurr = GetFileSize(strPathCurr)
	intSizePrev = GetFileSize(strPathPrev)
	
	If intSizeCurr <> intSizePrev Then
		WScript.Echo "Modification in group " & strGroupCn & " detected"
		
		Call SendMail(strGroupCn, strPathCurr, strPathPrev)
	Else	
		WScript.Echo "No changes in " & strGroupCn & " detected"
	End If
End Sub 



''gstrGroupDn = "CN=RP_SmartXS_Autorisatie,OU=Global Groups,OU=Beheer,DC=prod,DC=ns,DC=nl"
''Call CheckGroupForChanges(gstrGroupDn)
Call CheckGroupForChanges("DC=prod,DC=ns,DC=nl", "RP_SmartXS_Autorisatie")

WScript.Quit(0)



'' EOS