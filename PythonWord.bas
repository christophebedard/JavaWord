Function IsBinaryAccessible(binaryName As String) As Boolean
    Dim wShell As Object
    Dim exitCode As Long
    
    Set wShell = CreateObject("WScript.Shell")
    
    ' 1. Run "cmd /c where [binary]"
    ' "cmd /c" ensures we capture the exit code correctly.
    ' The last parameter "True" makes VBA wait for the command to finish.
    exitCode = wShell.Run("cmd /c where " & binaryName, 0, True)
    
    ' 2. Interpret the result
    ' 0 = Found
    ' 1 = Not Found
    If exitCode = 0 Then
        IsBinaryAccessible = True
    Else
        IsBinaryAccessible = False
    End If
    
    Set wShell = Nothing
End Function



Sub PythonWord()
' PythonWord macro
' runs a Word file as Python Script
'
' @author christophebedard

Dim strDocName As String
Dim strDocNameExt As String
Dim intPos As Integer
Dim strPythonPath As String
Dim strCmd As String

Dim textContent As String
Dim filePath As String
Dim fso As Object
Dim fileStream As Object

' Set path to python bin
strPythonPath = ""

If strPythonPath = "" Then
    If Not IsBinaryAccessible("python") Then
        MsgBox "Python could not be found. Add it to the path or edit the macro and set the instalation path.", vbCritical
        End
    End If
End If

' Retrieve name of ActiveDocument
strDocPath = ActiveDocument.Path
strDocName = ActiveDocument.Name
intPos = InStrRev(strDocName, ".")
strDocName = Left(strDocName, intPos - 1)
strDocNameExt = strDocName & ".py"

Application.ScreenUpdating = False

' Test if Activedocument has previously been saved
If ActiveDocument.Path = "" Then
    ' If not previously saved
    MsgBox "The current document must be saved at least once."
    End
Else
    ' If previously saved, create a copy
    ' Set myCopy = Documents.Add(ActiveDocument.FullName)
    textContent = ActiveDocument.Content.Text
    
    ' --- START OF CLEANUP SECTION ---
    ' We replace specific UTF-8 characters with standard ASCII equivalents.
    ' ChrW() tells Word the specific "Unicode ID" of the bad character.
    
    ' Replace Smart Double Quotes ( “ and ” ) with Straight Quotes (")
    textContent = Replace(textContent, ChrW(8220), """")
    textContent = Replace(textContent, ChrW(8221), """")
    
    ' Replace Smart Single Quotes ( ‘ and ’ ) with Straight Apostrophes (')
    textContent = Replace(textContent, ChrW(8216), "'")
    textContent = Replace(textContent, ChrW(8217), "'")
    
    ' Replace Em-dashes (—) with two hyphens (--)
    textContent = Replace(textContent, ChrW(8212), "--")
    
    ' Replace Ellipsis (…) with three dots (...)
    textContent = Replace(textContent, ChrW(133), "...")
    ' --- END OF CLEANUP SECTION ---
    
    filePath = strDocPath & "\" & strDocNameExt
    ' Copy File Content
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileStream = fso.CreateTextFile(filePath, True) ' True = overwrite if exists
    fileStream.Write textContent
    fileStream.Close
    
    Set fileStream = Nothing
    Set fso = Nothing
    
End If

Application.ScreenUpdating = True

' Call command to
'   1- cd to current directory
'   2- Add python to PATH
'   3- Run .py file
'   4- Pause and exit cmd window
strCmd = "cmd.exe /S /K "
strCmd = strCmd & " CD /D " & strDocPath
strCmd = strCmd & " & set PATH=" & strPythonPath & ";%PATH%"
strCmd = strCmd & " & python " & strDocNameExt
strCmd = strCmd & " & pause & exit"
Call Shell(strCmd, vbNormalFocus)

End Sub

