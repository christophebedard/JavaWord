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

Dim textContent As String
Dim filePath As String
Dim fso As Object
Dim fileStream As Object
    
' Set path to java bin
strJavaPath = "C:\Program Files\Common Files\Oracle\Java\javapath"

' Retrieve name of ActiveDocument
strDocPath = ActiveDocument.Path
strDocName = ActiveDocument.Name
intPos = InStrRev(strDocName, ".")
strDocName = Left(strDocName, intPos - 1)
strDocNameExt = strDocName & ".java"

Application.ScreenUpdating = False

' Test if Activedocument has previously been saved
If ActiveDocument.Path = "" Then
    ' If not previously saved
    MsgBox "The current document must be saved at least once."
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
'   2- Add java to PATH
'   3- Compile .java file
'   4- Run .java file
strCmd = "cmd.exe /S /K "
strCmd = strCmd & " CD /D " & strDocPath
strCmd = strCmd & " & set PATH=" & strJavaPath & ";%PATH%"
strCmd = strCmd & " & javac " & strDocNameExt
strCmd = strCmd & " & java " & strDocName
strCmd = strCmd & " & pause & exit"
Call Shell(strCmd, vbNormalFocus)

End Sub
