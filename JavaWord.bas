Attribute VB_Name = "NewMacros"
Sub JavaWord()
' JavaWord macro
' runs a Word file as Java code
'
' @author christophebedard

Dim strDocName As String
Dim strDocNameExt As String
Dim intPos As Integer
Dim strJavaPath As String
Dim strCmd As String

' Set path to java bin
strJavaPath = "C:\Program Files\Java\jdk-9.0.4\bin"

' Retrieve name of ActiveDocument
strDocPath = ActiveDocument.Path
strDocName = ActiveDocument.Name
intPos = InStrRev(strDocName, ".")
strDocName = Left(strDocName, intPos - 1)
strDocNameExt = strDocName & ".java"

' Test if Activedocument has previously been saved
If ActiveDocument.Path = "" Then
    ' If not previously saved
    MsgBox "The current document must be saved at least once."
Else
    ' If previously saved, create a copy
    Set myCopy = Documents.Add(ActiveDocument.FullName)

    ' Force the file to be saved
    If myCopy.Saved = False Then myCopy.Save FileName:=strDocPath & "\" & strDocName

    ' Save file with new extension
    myCopy.SaveAs2 FileName:=strDocPath & "\" & strDocNameExt, FileFormat:=wdFormatText

    ' Close copy
    myCopy.Close
End If

' Call command to
'   1- cd to current directory
'   2- Add java to PATH
'   3- Compile .java file
'   4- Run .java file
strCmd = "cmd.exe /S /K"
strCmd = strCmd & "CD " & strDocPath
strCmd = strCmd & " & set PATH=" & strJavaPath & ";%PATH%"
strCmd = strCmd & " & javac " & strDocNameExt
strCmd = strCmd & " & java " & strDocName
Call Shell(strCmd, vbNormalFocus)

End Sub
