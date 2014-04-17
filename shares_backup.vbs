Const ForReading = 1
Const ForWriting = 2
Const bWaitOnReturn = True

' WMI Constants

Const WBEM_RETURN_IMMEDIATELY = &h10
Const WBEM_FORWARD_ONLY = &h20

' Constants and storage arrays for security settings

' GetSecurityDescriptor Return values

Dim objReturnCodes : Set objReturnCodes = CreateObject("Scripting.Dictionary")
Const SUCCESS = 0
Const ACCESS_DENIED = 2
Const UNKNOWN_FAILURE = 8
Const PRIVILEGE_MISSING = 9
Const INVALID_PARAMETER = 21

' Security Descriptor Control Flags

Dim objControlFlags : Set objControlFlags = CreateObject("Scripting.Dictionary")
objControlFlags.Add 32768, "SelfRelative"
objControlFlags.Add 16384, "RMControlValid"
objControlFlags.Add 8192, "SystemAclProtected"
objControlFlags.Add 4096, "DiscretionaryAclProtected"
objControlFlags.Add 2048, "SystemAclAutoInherited"
objControlFlags.Add 1024, "DiscretionaryAclAutoInherited"
objControlFlags.Add 512, "SystemAclAutoInheritRequired"
objControlFlags.Add 256, "DiscretionaryAclAutoInheritRequired"
objControlFlags.Add 32, "SystemAclDefaulted"
objControlFlags.Add 16, "SystemAclPresent"
objControlFlags.Add 8, "DiscretionaryAclDefaulted"
objControlFlags.Add 4, "DiscretionaryAclPresent"
objControlFlags.Add 2, "GroupDefaulted"
objControlFlags.Add 1, "OwnerDefaulted"

' ACE Access Right

Dim objAccessRights : Set objAccessRights = CreateObject("Scripting.Dictionary")
objAccessRights.Add 2032127, "FullControl"
objAccessRights.Add 1048576, "Synchronize"
objAccessRights.Add 524288, "TakeOwnership"
objAccessRights.Add 262144, "ChangePermissions"
objAccessRights.Add 197055, "Modify"
objAccessRights.Add 131241, "ReadAndExecute"
objAccessRights.Add 131209, "Read"
objAccessRights.Add 131072, "ReadPermissions"
objAccessRights.Add 65536, "Delete"
objAccessRights.Add 278, "Write"
objAccessRights.Add 256, "WriteAttributes"
objAccessRights.Add 128, "ReadAttributes"
objAccessRights.Add 64, "DeleteSubdirectoriesAndFiles"
objAccessRights.Add 32, "ExecuteFile"
objAccessRights.Add 16, "WriteExtendedAttributes"
objAccessRights.Add 8, "ReadExtendedAttributes"
objAccessRights.Add 4, "AppendData"
objAccessRights.Add 2, "CreateFiles"
objAccessRights.Add 1, "ReadData"

' ACE Types

Dim objAceTypes : Set objAceTypes = CreateObject("Scripting.Dictionary")
objAceTypes.Add 0, "Allow"
objAceTypes.Add 1, "Deny"
objAceTypes.Add 2, "Audit"

' ACE Flags

Dim objAceFlags : Set objAceFlags = CreateObject("Scripting.Dictionary")
objAceFlags.Add 128, "FailedAccess"
objAceFlags.Add 64, "SuccessfulAccess"
objAceFlags.Add 16, "Inherited"
objAceFlags.Add 8, "InheritOnly"
objAceFlags.Add 4, "NoPropagateInherit"
objAceFlags.Add 2, "ContainerInherit"
objAceFlags.Add 1, "ObjectInherit"

Sub ReadNTFSSecurity(objWMI, strPath)
  objFileOut.Write("  Displaying NTFS Security" & vbCrLf)

  Dim objSecuritySettings : Set objSecuritySettings = _
    objWMI.Get("Win32_LogicalFileSecuritySetting='" & strPath & "'")
  Dim objSD : objSecuritySettings.GetSecurityDescriptor objSD

  Dim strDomain : strDomain = objSD.Owner.Domain
  If strDomain <> "" Then strDomain = strDomain & "\"
  objFileOut.Write("  Owner: " & strDomain & objSD.Owner.Name & vbCrLf)
  objFileOut.Write("  Owner SID: " & objSD.Owner.SIDString & vbCrLf)

  objFileOut.Write("  Basic Control Flags Value: " & objSD.ControlFlags & vbCrLf)
  objFileOut.Write("  Control Flags:" & vbCrLf)

  DisplayValues objSD.ControlFlags, objControlFlags

  objFileOut.Write(vbCrLf)

  Dim objACE

  ' Display the DACL
  objFileOut.Write("  Discretionary Access Control List:" & vbCrLf)
  For Each objACE in objSD.DACL
    DisplayACE objACE
  Next

  ' Display the SACL (if there is one)
  If Not IsNull(objSD.SACL) Then
    objFileOut.Write("  System Access Control List:" & vbCrLf)
    For Each objACE in objSD.SACL
      DisplayACE objACE
    Next
  End If
End Sub

Sub ReadShareSecurity(objWMI, strName)

	Dim objSecuritySettings : Set objSecuritySettings = objWMI.Get("Win32_LogicalShareSecuritySetting='" & strName & "'")
	Dim objSD : objSecuritySettings.GetSecurityDescriptor objSD
	Dim objACE

	For Each objACE in objSD.DACL
		DisplayACE objACE
	Next

End Sub

Sub DisplayValues(dblValues, objSecurityEnumeration)

	Dim dblValue

	For Each dblValue in objSecurityEnumeration
		If dblValues >= dblValue Then
			If (objSecurityEnumeration(dblValue) <> "Synchronize") Then objFileOut.Write("Right:" & objSecurityEnumeration(dblValue) & vbCrLf)
			dblValues = dblValues - dblValue
	    End If
	Next

End Sub

Sub DisplayACE(objACE)

	Dim strDomain : strDomain = objAce.Trustee.Domain

	If strDomain <> "" Then strDomain = strDomain & "\"

	If (UCase(objAceTypes(objACE.AceType)) = "ALLOW") Then
		objFileOut.Write("Trustee:" & UCase(strDomain & objAce.Trustee.Name) & vbCrLf)
		DisplayValues objACE.AccessMask, objAccessRights
	End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Main Code
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next

Set oShell = CreateObject("WScript.Shell")
SysDrive=oShell.ExpandEnvironmentStrings("%SystemDrive%")
ProcArch=oShell.ExpandEnvironmentStrings("%processor_architecture%")
Set objFSO = CreateObject("Scripting.FileSystemObject")
CurrentDirectory = objFSO.GetAbsolutePathName(".")
Set objFolder = objFSO.CreateFolder(CurrentDirectory & "\")
Set objFolder = Nothing
Set colDrives = objFSO.Drives
Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True   
objRegEx.IgnoreCase = True
objRegEx.Pattern = ".*\\"
Set objREx = CreateObject("VBScript.RegExp")
objREx.Global = True   
objREx.IgnoreCase = True
objREx.Pattern = "[\:\\ ]"
Set objFileOut = objFSO.OpenTextFile(CurrentDirectory & "\shares.txt", ForWriting, True)
Set objFileOut2 = objFSO.OpenTextFile(CurrentDirectory & "\shares_copy.cmd", ForWriting, True)
Set objFileOut3 = objFSO.OpenTextFile(CurrentDirectory & "\shares_create.cmd", ForWriting, True)
Set objFileOut6 = objFSO.OpenTextFile(CurrentDirectory & "\own_shares.cmd", ForWriting, True)
objFileOut2.Write("@echo off" & vbCrLf)
objFileOut2.Write("chcp 1251" & vbCrLf)
objFileOut3.Write("@echo off" & vbCrLf)
objFileOut3.Write("set scriptpath=%~dp0" & vbCrLf)
objFileOut3.Write(vbCrLf & "IF %processor_architecture% == x86 (" & vbCrLf & "set setacl=SetACL32.exe" & vbCrLf & ") ELSE (" & vbCrLf & "set setacl=SetACL64.exe" & vbCrLf & ")" & vbCrLf & vbCrLf)
objFileOut6.Write("@echo off" & vbCrLf)
objFileOut6.Write("chcp 1251" & vbCrLf)
objFileOut6.Write("set scriptpath=%~dp0" & vbCrLf)
objFileOut6.Write(vbCrLf & "IF %processor_architecture% == x86 (" & vbCrLf & "set setacl=SetACL32.exe" & vbCrLf & ") ELSE (" & vbCrLf & "set setacl=SetACL64.exe" & vbCrLf & ")" & vbCrLf & vbCrLf)

Dim strComputer : strComputer = "."
Dim objWMI : Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Dim colItems : Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_Share WHERE Type='0'", "WQL", WBEM_RETURN_IMMEDIATELY + WBEM_FORWARD_ONLY)
Dim objItem

If (ProcArch = "x86") Then
	setacl = "SetACL32.exe"
Else
	setacl = "SetACL64.exe"
End If

dircounter = 0

copycommand = "copy /A /Y "

'Creating container folder for existence flags of shared folders (protection from folders with more than 1 shares)
If (Not objFSO.FolderExists(CurrentDirectory & "\flags")) Then
	oShell.run "cmd /c mkdir """ & CurrentDirectory & "\flags""",0,bWaitOnReturn
End If

For Each objItem in colItems

		objFileOut.Write("Share:" & objItem.Name & vbCrLf)
		objFileOut.Write("Path:" & objItem.Path & vbCrLf)
		objFileOut.Write("Desc:" & objItem.Caption & vbCrLf)
		ReadShareSecurity objWMI, objItem.Name
		'ReadNTFSSecurity objWMI, objItem.Path
		objFileOut.Write(vbCrLf)
		
		If Not objFSO.FileExists(CurrentDirectory & "\flags\" & objREx.Replace(objItem.Path,"_")) Then
			'Generating shares_copy.cmd
			objFileOut2.Write("xcopy /y /k /e /z """ & objItem.Path & "\*"" """ & "%1\" & objRegEx.Replace(objItem.Path,"") & "\*""" & vbCrLf)
			objFileOut2.Write("if not %errorlevel%==0 exit %errorlevel%" & vbCrLf)
			
			'Generating own_shares.cmd
			objFileOut6.Write("echo Taking ownership on " & objItem.Path & " ..." & vbCrLf)
			objFileOut6.Write("takeown /R /A /D ""Y"" /F """ & objItem.Path & """ > NUL" & vbCrLf)
			objFileOut6.Write("if not %errorlevel%==0 (" & vbCrLf & "	echo Failed." & vbCrLf & ") else (" & vbCrLf & "	echo Done." & vbCrLf & ")" & vbCrLf)
			objFileOut6.Write("echo Adding Administrators and SYSTEM to ACL on " & objItem.Path & " ..." & vbCrLf)
			objFileOut6.Write("""%scriptpath%%setacl%"" -silent -ot file -on """ & objItem.Path & """ -actn ace -ace ""n:S-1-5-18;p:full"" -ace ""n:S-1-5-32-544;p:full""" & vbCrLf)
			objFileOut6.Write("if not %errorlevel%==0 (" & vbCrLf & "	echo Failed." & vbCrLf & ") else (" & vbCrLf & "	echo Done." & vbCrLf & ")" & vbCrLf)
			
			'Making backup of NTFS ACL for shared folders
			oShell.Run """" & CurrentDirectory & "\" & setacl & """ -on """ & objItem.Path & """ -ot file -actn list -lst ""f:sddl;w:d,s,o,g"" -bckp """ & CurrentDirectory & "\" & dircounter & ".acl""",0,bWaitOnReturn
			If (dircounter = 0) Then
				copycommand = copycommand & """" & CurrentDirectory & "\" & dircounter & ".acl"""
			Else
				copycommand = copycommand & "+""" & CurrentDirectory & "\" & dircounter & ".acl"""
			End If
			dircounter = dircounter + 1
			
			'Creating flag for folder which already was processed
			Set objFlagObj = objFSO.OpenTextFile(CurrentDirectory & "\flags\" & objREx.Replace(objItem.Path,"_"), ForWriting, True)
			objFlagObj.Close
		End If
	
Next

'Removing container folder for flags
If (objFSO.FolderExists(CurrentDirectory & "\flags")) Then
	oShell.run "cmd /c rd /s /q """ & CurrentDirectory & "\flags""",0,bWaitOnReturn
End If

'Merging files with ACL for every shared folder
copycommand = copycommand & " """ & CurrentDirectory & "\acllist.lca"""
oShell.run "cmd /c " & copycommand,0,bWaitOnReturn
oShell.run "cmd /c del /F /Q """ & CurrentDirectory & "\*.acl""",0,bWaitOnReturn

'Changing codepage for file with ACL from UCS-2 LE (UTF-16) to UTF-8
Set ADODBStream = CreateObject("ADODB.Stream")
ADODBStream.Type = 2
ADODBStream.Charset = "UTF-16LE"
ADODBStream.Open()
ADODBStream.LoadFromFile(CurrentDirectory & "\acllist.lca")
Text = ADODBStream.ReadText()
ADODBStream.Close()
ADODBStream.Charset = "UTF-8"
ADODBStream.Open()
ADODBStream.WriteText(Text)
ADODBStream.SaveToFile CurrentDirectory & "\acllist.lca", 2
ADODBStream.Close()

objFileOut3.Write("cscript.exe ""%scriptpath%shares_restore.vbs""" & vbCrLf)
objFileOut3.Write("""%scriptpath%%setacl%"" -ignoreerr -on ""%SystemDrive%"" -ot file -actn restore -bckp ""%scriptpath%acllist.lca""" & vbCrLf)
objFileOut.Close
objFileOut2.Close
objFileOut3.Close
objFileOut6.Close
