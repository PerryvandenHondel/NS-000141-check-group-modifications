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
''


Option Explicit




Dim	gstrGroupDn




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



Sub CheckGroupForChanges(ByVal strRootDse, ByVal strGroupCn)
	Dim		c
	Dim		strPathCurrent
	Dim		strPathLast
	
	WScript.Echo "Check group: " & strGroupCn

	strPathCurrent = strGroupCn & "-current.txt"
	strPathLast = strGroupCn & "-last.txt"
	
	WScript.Echo strPathCurrent
	WScript.Echo strPathLast
	
	'' file current.txt exist, rename to last.txt
	
	'' 
	c = "adfind.exe -b " & EncloseWithDQ(strRootDse) & " -f " & EncloseWithDQ("CN=" & strGroupCn) & " member -list >" & strPathCurrent
	WScript.Echo c
End Sub



''gstrGroupDn = "CN=RP_SmartXS_Autorisatie,OU=Global Groups,OU=Beheer,DC=prod,DC=ns,DC=nl"
''Call CheckGroupForChanges(gstrGroupDn)
Call CheckGroupForChanges("DC=prod,DC=ns,DC=nl", "RP_SmartXS_Autorisatie")

WScript.Quit(0)



'' EOS