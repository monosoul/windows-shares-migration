Const ForReading = 1
Const ForWriting = 2

Const FILE_SHARE          = 0 
Const MAXIMUM_CONNECTIONS = 4294967295 

On Error Resume Next
Set oShell = CreateObject("WScript.Shell")
SysDrive=oShell.ExpandEnvironmentStrings("%SystemDrive%")
ProcArch=oShell.ExpandEnvironmentStrings("%processor_architecture%")
Set WSHShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
PathToScript = Wscript.ScriptFullName
Set ScriptFile = objFSO.GetFile(PathToScript)
CurrentDirectory = objFSO.GetParentFolderName(ScriptFile)
Set objFileIn = objFSO.OpenTextFile(CurrentDirectory & "\shares.txt", ForReading)
Set colDrives = objFSO.Drives

If (ProcArch = "x86") Then
	setacl = "SetACL32.exe"
Else
	setacl = "SetACL64.exe"
End If

Do Until objFileIn.AtEndOfStream
	ProcessShare()
Loop

objFileIn.Close

Sub ProcessShare()

share_name = ""
share_path = ""
share_desc = ""

Do Until objFileIn.AtEndOfStream Or ((Len(share_name) > 0) And (Len(share_path) > 0) And (Len(share_desc) > 0))
	tmpline = objFileIn.ReadLine
	If (Len(tmpline) > 0) Then
		If (InStr(tmpline, "Share:") = 1) Then share_name = Right(tmpline, Len(tmpline) - Len("Share:"))
		If (InStr(tmpline, "Path:") = 1) Then share_path = Right(tmpline, Len(tmpline) - Len("Path:"))
		If (InStr(tmpline, "Desc:") = 1) Then share_desc = Right(tmpline, Len(tmpline) - Len("Desc:"))
	End If
Loop 'Read all parameters for share creation, creating share

If ((Len(share_name) > 0) And (Len(share_path) > 0) And (Len(share_desc) > 0)) Then

	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set objNewShare = objWMIService.Get("Win32_Share")
	WScript.Echo share_name & " (" & share_path & "):"
	errReturn = objNewShare.Create (share_path, share_name, FILE_SHARE, MAXIMUM_CONNECTIONS, share_desc)

	If errReturn = 0 Then
		'Created share, removing default trustees
		WSHShell.Run """" & CurrentDirectory & "\" & setacl & """ -ot shr -on """ & share_name & """ -actn trustee -trst ""n1:""ВСЕ"";ta:remtrst""", 1, True
		WSHShell.Run """" & CurrentDirectory & "\" & setacl & """ -ot shr -on """ & share_name & """ -actn trustee -trst ""n1:""EVERYONE"";ta:remtrst""", 1, True

		Do
			trustee = ""
			perm = ""
			Do Until objFileIn.AtEndOfStream Or ((Len(trustee) > 0) And (Len(perm) > 0)) Or (Len(tmpline) < 2)
				tmpline = objFileIn.ReadLine
				If (Len(tmpline) > 0) Then
					If (InStr(tmpline, "Trustee:") = 1) Then trustee = Right(tmpline, Len(tmpline) - Len("Trustee:"))
					If (InStr(tmpline, "Right:") = 1) Then perm = Right(tmpline, Len(tmpline) - Len("Right:"))
				End If
			Loop 'Read trustee and rights, applying

			If ((Len(trustee) > 0) And (Len(perm) > 0)) Then
				If (InStr(UCase(trustee), "BUILTIN\") = 1) Then trustee = Right(trustee, Len(trustee) - Len("BUILTIN\"))
				If (perm = "FullControl") Then perm = "full"
				If (perm = "ReadAndExecute") Then perm = "read"
				If (perm = "Modify") Then perm = "change"
				WScript.Echo trustee & " - " & perm 
				WSHShell.Run """" & CurrentDirectory & "\" & setacl & """ -ot shr -on """ & share_name & """ -actn ace -ace ""n:" & trustee & ";p:" & perm & ";m:set""", 1, True
			End If
		Loop Until (Len(tmpline) < 2) Or objFileIn.AtEndOfStream
	Else
		Select Case errReturn
			Case 2
				errText = "Access Denied"
			Case 8
				errText = "Unknown Problem"
			Case 9
				errText = "Invalid Name"
			Case 10
				errText = "Invalid Level"
			Case 21
				errText = "Invalid Parm"
			Case 22
				errText = "Share Already Exists"
			Case 23
				errText = "Redirected Path"
			Case 24
				errText = "Missing Folder"
			Case 25
				errText = "Missing Server"
			Case Else
				errText = "Operation could not be completed"
		End Select
		WScript.Echo "Error while creating shared folder " & share_name & "! Error code:" & errReturn & " - " & errText
	End If

End If

End Sub
