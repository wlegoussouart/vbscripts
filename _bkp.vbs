' ======================================================================
' BEGIN VBSCRIPT
' ======================================================================
' dd/mm/yyyy author          	description
' ---------- ------------------	----------------------------------------
' 16/07/2007 jgantner           created
' 08/07/2009 wlegoussouart      1- Use the Last Updated Date of the file 
'                                   for label instead of Now().
'                               2- Added a msg box to display which files 
'                                   were already existing in backup folder.
' 25/05/2011 wlegoussouart      1- Use the name of the script as backup 
'                                   folder name.
' 18/05/2012 jgantner           lower-case bitwise "and" does not appear
'                               to detect existing RO attribute (1) resulting 
'                               in existing RO file backed up with hidden flag set (1+1=2)!
'
' 06/08/2014 jgantner			Modularised the script to leave only the "Call Main()" 
'								call in the body.
'
' 14/06/2016 wlegoussouart		1- Optionally resolve shortcuts to save a 
'									copy of the target file in the local 
'									backup folder. 
'								2- Changed all the msgboxes to be WScript.Echo
'									, this is more cscript friendly
'									, in case this is run in console mode. 
'								<TODO>: accept wildcards, like *.vbs...
'								<TODO>: if running in cscript, we can afford to be more verbose.
'								<TODO>: if running in cscript in console, we should be able 
'									to create switches (eg /ResolveShortcutFile, name of the backup folder...).
' ----------------------------------------------------------------------
Option Explicit

' Global Parameters
' ----------------------------------------------------------------------
Const ResolveShortcutFile = true

'Call the main routine!
Call Main()


'* MAIN ======================================================================
Sub Main
    Dim arrPaths, path
	Dim FinalMessage, RetMsg
    arrPaths = GetFilePaths
	
	'Init the Final message
	'(we want only one message at the end, not one for each already existing file, for example.)
	FinalMessage = ""

    If UBound(arrPaths) < 0 Then
        'FinalMessage = "Usage:" & vbcrlf & vbcrlf & Wscript.ScriptName & " Filename"
		FinalMessage = "Usage:" & vbcrlf & vbcrlf & _
						"" & Wscript.ScriptName & " Filename1[ Filename2[ Filename3...]]" & vbcrlf & vbcrlf & _ 
						" Or " & vbcrlf & vbcrlf & _
						"cscript " & Wscript.ScriptName & " Filename1[ Filename2[ Filename3...]]"
    Else
        For each path in arrPaths
			'Backup the file
			RetMsg = BackupFile(path)
			'Append the message (if there is one)
			If Len(Trim(RetMsg))> 0 Then
				FinalMessage = FinalMessage & vbCrLf & RetMsg
			End If
        Next
		

	End If
	'Display the final Message.
	If Len(Trim(FinalMessage))> 0 Then
		'Call MsgBox(FinalMessage, vbOkOnly+vbInformation, "Backup the files")
		Wscript.Echo FinalMessage
	End If
End Sub

'* ROUTINES ==================================================================
Function BackupFile(pathSrc)
    Const ReadOnly = 1
    Dim fso, f, dtm
    Dim nameDestFile, pathDestFolder, pathDest, extn
	Dim pathResolvedSrc
	
    Set fso = WScript.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(pathSrc) Then
        Set f = fso.GetFile(pathSrc)
        dtm = f.DateLastModified
    Else
        ' file does not exist!
		'call DisplayError("The file """& pathSrc & """ does not exist!")
		BackupFile = "The file """& pathSrc & """ does not exist!"
        Exit Function
    End If
	
	'Get the absolute path, sometimes (using cscript, for example, the context is lost)
	pathSrc = fso.GetAbsolutePathName(pathSrc)
	
	'Creating the destination folder, which will have the same name as the script filename (backup, _bkp, or whatever else the user renamed it to.)
    pathDestFolder = fso.GetParentFolderName(pathSrc) & "\" & fso.GetBaseName(WScript.ScriptName)
    If Not fso.FolderExists(pathDestFolder) Then
        fso.CreateFolder pathDestFolder
    End If
	
	'<debug>
	'Wscript.Echo "fso.GetParentFolderName(pathSrc) = " & fso.GetParentFolderName(pathSrc)
	'Wscript.Echo "pathSrc = " & pathSrc
	'Wscript.Echo "pathDestFolder = " & pathDestFolder
	
	'Check that the backup folder was indeed created.
    If Not fso.FolderExists(pathDestFolder) Then
        ' backup folder could not be created!
		BackupFile = "The backup Folder """& pathDestFolder & """ could not be created!"
		'call DisplayError("The backup Folder """& pathDestFolder & """ could not be created!")
        Exit Function
    End If

    extn = fso.GetExtensionName(pathSrc)
    If Len(extn) > 0 Then extn = "."  & extn
	
	'Resolve the shortcut if needed.
	If ResolveShortcutFile And (extn = ".lnk" Or extn = ".url")  Then
		'resolve the shortcut if pathsrc is a shortcut
		pathResolvedSrc = GetShortcut(pathSrc)
		
		'<debug>
		'Wscript.Echo "Resolved Path Source: " & pathResolvedSrc

		pathSrc = pathResolvedSrc
		'we need to get the extn again, it is no longer .lnk or .url :)
		extn = fso.GetExtensionName(pathSrc)
		If Len(extn) > 0 Then extn = "."  & extn
	End If
	
    nameDestFile = fso.GetBaseName(pathSrc) & "_" & GetDateStamp(dtm) & extn
    pathDest = pathDestFolder & "\" & nameDestFile

	'Check if the file already exists in target location. 
	'This is actually a  way to verify it has been created correctly ;)
    If fso.FileExists(pathDest) Then
		BackupFile = "File " & fso.GetFileName(pathdest) & " already exists."
        'WScript.Echo "file already exists: " & pathDest
		'MsgBox "file already exists: " & pathDest
        Exit Function
    Else
        fso.CopyFile pathSrc, pathDest, false
		Call EnsureReadOnly(pathDest)
		BackupFile=""
		'<TODO> we might want to send back an error if an error occurs.
    End If

End Function

'-----------------------------------------------------------------------------
Sub DisplayError(msg)
    'MsgBox msg, vbExclamation, Wscript.ScriptName
	Wscript.Echo msg
End Sub

'-----------------------------------------------------------------------------
Function GetFilePaths
    Const cLIST_DELIM = "|"

    Dim objArgs, arg
    Dim strPaths
	Dim strMsg

    Set objArgs = WScript.Arguments

    strPaths = ""
    If objArgs.Count > 0 Then

        For Each arg in objArgs
            strPaths = strPaths & CStr(arg) & cLIST_DELIM
        Next

        If Len(strPaths) > 0 Then
            ' strip trailing delimiter
            strPaths = Left(strPaths, Len(strPaths) - len(cLIST_DELIM))
        End If

        GetFilePaths = Split(strPaths, cLIST_DELIM, -1, vbTextCompare)	'-1: all substrings are returned
    Else
        GetFilePaths = Array()
    End if
	
	'<debug>
	'strMsg = "The paths that were found: "& strPaths
	'MsgBox strMsg, vbInformation, "debug"
	'Wscript.Echo strMsg

End Function

'-----------------------------------------------------------------------------
Function GetDateStamp(dtm)

    On Error Resume Next

    GetDateStamp = _
		CStr(Year(dtm)) & _
		LZ(Month(dtm)) & _
		LZ(Day(dtm)) & _
		"_" & _
		LZ(Hour(dtm)) & _
		LZ(Minute(dtm)) & _
		LZ(Second(dtm))

End Function

'-----------------------------------------------------------------------------
Function LZ(ByVal number)

    If number < 10 Then
        LZ = "0" & CStr(number)
    Else
        LZ = CStr(number)
    End If

End Function

'-----------------------------------------------------------------------------
Sub EnsureReadOnly(filespec)
	Dim fso, f
	Const ReadOnly = 1
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFile(filespec)
	'Need bitwise "AND" to be uppercase to be BITWISE!
	If Not (f.attributes AND ReadOnly) Then
		f.attributes = f.attributes + ReadOnly
	End If
	Set f = Nothing
	Set fso = Nothing
End Sub


'-----------------------------------------------------------------------------
Function GetShortcut(tgtPath)
	' With the help from source: http://www.robvanderwoude.com/vbstech_shortcuts.php (Author: Denis St-Pierre)
	' *Retrieves* Shortcut info without using WMI 
	' The *Undocumented* Trick: use the ".CreateShortcut" method without the 
	' ".Save" method; works like a GetShortcut when the shortcut already exists!
	Dim wshShell, objShortcut
	
	Set wshShell = CreateObject("WScript.Shell")
	' CreateShortcut works like a GetShortcut when the shortcut already exists!

	Set objShortcut = wshShell.CreateShortcut(tgtPath)
	If len(trim(objShortcut.TargetPath))>0 then 
		GetShortcut = objShortcut.TargetPath	
	Else
		GetShortcut = ""
	End If
	
	'<debug>
	' Note: for URL shortcuts, only ".FullName" and ".TargetPath" are valid
	'WScript.Echo "Full Name         : " & objShortcut.FullName
	'WScript.Echo "Arguments         : " & objShortcut.Arguments
	'WScript.Echo "Working Directory : " & objShortcut.WorkingDirectory
	'WScript.Echo "Target Path       : " & objShortcut.TargetPath
	'WScript.Echo "Icon Location     : " & objShortcut.IconLocation
	'WScript.Echo "Hotkey            : " & objShortcut.Hotkey
	'WScript.Echo "Window Style      : " & objShortcut.WindowStyle
	'WScript.Echo "Description       : " & objShortcut.Description

	Set objShortcut = Nothing
	Set wshShell    = Nothing
End Function