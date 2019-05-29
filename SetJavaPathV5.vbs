' v5.0  August 21, 2010
' Antonio C. Silvestri
' USAGE: cscript|wscript SetJavaPath5.vbs
' DESCRIPTION:  Searches for the first instance of JAVAC.EXE using FSO
' There should only be one JDK installed anyway.  Taking this into account,
' this version is much faster in setting the user environment path to 
' the directory in which the file is found
'***********************************************************************************

Option Explicit
On error Resume Next

Dim objProgressMsg

Const ProgramTitle = "JAVA PATH Environment Setter V5.0"
Const strName = "JAVAC.EXE"

Call SetJavaPath()
wscript.quit

'***********************************************************************************

Sub SetJavaPath()
  On Error Resume Next
  Dim JavaPath
  Dim Prompt

  prompt = "Sets Your PATH Environment Variable to point to your Java Installation." & vbNewLine & _
           "(It Could Take Awhile to Find Your JDK.  So Be Patient!!!)"
  ProgressMsg prompt, ProgramTitle, True

  JavaPath = SearchForFirst("C:\Program Files", strName)
  If JavaPath = "" Then
    Prompt = "Cannot Locate JAVAC.EXE.  Perhaps the JDK was never installed."
    ProgressMsg prompt, ProgramTitle, False
    wscript.quit
  End If

  SetEnvironmentVariable JavaPath

  If Err.Number = 0 Then
    prompt = "Created USER PATH Environment Variable: " & JavaPath & vbNewLine
  Else
    prompt = "Error Occurred: " & Err.Number & " " & Err.Description & vbNewLine
  End If
  ProgressMsg prompt, ProgramTitle, False
End Sub

'***********************************************************************************
' Test Code for return ALL Directories where JDK Is Found
' 
' Seach is not used in this program
'
' Dim paths, path
' Set paths = search("C:\Program Files")
' WScript.Echo paths.count
' For Each path In paths.keys
'   WScript.Echo path
' next
' wscript.quit

Function Search(ByVal directory, ByVal strName)
  On Error Resume Next
  Dim objFSO, currentFolder, objFile
  Dim objFolder, files, tempFiles, path
  Dim currFiles, currFolds

  Set Search = Nothing
  Set files = WScript.CreateObject("Scripting.Dictionary")
  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
  Set currentFolder = objFSO.GetFolder(directory)
  Set currFiles = currentFolder.Files
  Set currFolds = currentFolder.SubFolders
  For Each objFile In currFiles
 	If UCase(objFile.Name) = strName Then
      files.Add directory, directory
	  Exit For
	End If
  Next
  For Each objFolder in currFolds
	Set tempFiles = Search( objFolder, strName )
	If Not (tempFiles Is Nothing) Then 
	  For Each path In tempFiles.keys
	    files.Add path, path
	  Next
	  Set tempFiles = Nothing
	End if
  Next
  Set currentFolder = Nothing
  Set currFiles = Nothing
  Set currFolds = Nothing
  Set objFSO = Nothing
  Set Search = files
  Set files = Nothing
End Function

'***********************************************************************************

Function SearchForFirst(ByVal directory, ByVal strName)
  On Error Resume Next
  Dim objFSO, currentFolder, objFile
  Dim objFolder, tempFiles, path
  Dim currFiles, currFolds

  SearchForFirst = ""
  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
  Set currentFolder = objFSO.GetFolder(directory)
  Set currFiles = currentFolder.Files
  For Each objFile In currFiles
 	If UCase(objFile.Name) = strName Then
      SearchForFirst = directory
	  Exit For
	End If
  Next
  If SearchForFirst = "" Then
    Set currFolds = currentFolder.SubFolders
    For Each objFolder in currFolds
	  tempFiles = SearchForFirst( objFolder, strName )
	  If tempFiles <> "" Then
	    SearchForFirst = tempFiles
	    Exit For
      End if
    Next
    Set currFolds = Nothing
  End if
  Set currFiles = Nothing
  Set currentFolder = Nothing
  Set objFSO = Nothing
End Function

'***********************************************************************************
' http://www.robvanderwoude.com/vbstech_data_environment.php

Sub SetEnvironmentVariable(ByVal JavaPath)
  Dim wshShell, wshSystemEnv
  Dim Dirs
  Dim PathValue, NewPath
  Dim i
  
  Set wshShell = WScript.CreateObject( "WScript.Shell" )
  Set wshSystemEnv = wshShell.Environment( "USER" )

  PathValue = wshSystemEnv( "PATH" )
  Dirs = Split(PathValue, ";", -1, 1)
  NewPath = ""
  For i = 0 To UBound(Dirs)
    If InStr(1, Dirs(i), "JAVA", 1) = 0 And InStr(1, Dirs(i), "JDK", 1) = 0 Then
      NewPath = NewPath & Dirs(i)
	  If i < UBound(Dirs) Then
	    NewPath = NewPath & ";"
	  End If
    End if
  Next
  If NewPath <> "" Then
    NewPath = ";" & NewPath
  End if
  NewPath = JavaPath & NewPath

  SetPathEnvironment NewPath

  Set wshSystemEnv = Nothing
  Set wshShell     = Nothing
End Sub

'***********************************************************************************  

Sub SetPathEnvironment(ByVal NewPath)
  Dim wshShell, wshSystemEnv
  Set wshShell = WScript.CreateObject( "WScript.Shell" )
  Set wshSystemEnv = wshShell.Environment( "USER" )
  wshSystemEnv( "PATH" ) = NewPath
  Set wshSystemEnv = Nothing
  Set wshShell     = Nothing
End Sub

'***********************************************************************************  

Function ProgressMsg(ByVal strMessage, ByVal strWindowTitle, ByVal Force )
' http://www.robvanderwoude.com/vbstech_ui_progress.php
' If StrMessage is blank, take down previous progress message box
' Using 4096 in Msgbox below makes the progress message float on top of things
' CAVEAT: You must have   Dim ObjProgressMsg   at the top of your script for this to work as described
  Dim wshShell
  Dim strTemp
  Dim objFSO
  Dim strTempVBS
  Dim objTempMessage
  Dim PromptLines
  Dim i

  Set wshShell = WScript.CreateObject( "WScript.Shell" )
  strTEMP = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
  If strMessage = "" Then
    ' Disable Error Checking in case objProgressMsg doesn't exists yet
    On Error Resume Next
    ' Kill ProgressMsg
    objProgressMsg.Terminate()
    ' Re-enable Error Checking
    On Error Goto 0
  else
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strTempVBS = strTEMP + "\" & "Message.vbs"
    Set objTempMessage = objFSO.CreateTextFile( strTempVBS, True )
    objTempMessage.WriteLine("Dim Prompt")
    objTempMessage.WriteLine("Prompt = """"")
    PromptLines = Split(strMessage, vbNewLine)
    For i = 0 To UBound(PromptLines)
      objTempMessage.Write( "Prompt = Prompt & """ & PromptLines(i) & """" )
      If i < UBound(PromptLines) Then
	    objTempMessage.Write (" & vbNewLine ")
	  End If
	  objTempMessage.WriteLine()
    Next 
	If Force Then
	  objTempMessage.WriteLine("Do While True")
	End if
    objTempMessage.WriteLine("MsgBox Prompt, 4096, """ & strWindowTitle & """")
	If Force Then
	  objTempMessage.WriteLine("Loop")
	End if
    objTempMessage.Close
    Set objFSO   = Nothing

    ' Disable Error Checking in case objProgressMsg doesn't exists yet
    On Error Resume Next
    ' Kills the Previous ProgressMsg
    objProgressMsg.Terminate( )
    ' Re-enable Error Checking
    On Error Goto 0
    ' Trigger objProgressMsg and keep an object on it
    Set objProgressMsg = WshShell.Exec( "%windir%\system32\wscript.exe " & strTempVBS )
  End if
  Set wshShell = Nothing
End Function
