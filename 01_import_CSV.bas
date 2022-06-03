Option Explicit

Sub importCSV()

Dim objFileSystemObject As Object
Dim objLine As Object
Dim strLine, strFilename As String
Dim intCounter As Integer
Dim wksSheet As Worksheet

Set wksSheet = Sheet1     ' Name of an existing worksheet within the current workbook.
strFilename = "Input.csv" ' Define the filename of the import file.

wksSheet.Rows.Delete
Set objFileSystemObject = CreateObject("Scripting.FilesystemObject")
Set objLine = objFileSystemObject.OpenTextFile(ThisWorkbook.Path & _
    "\" & strFilename)

With wksSheet
    intCounter = 1
    Do Until objLine.AtEndOfStream
        strLine = objLine.Readline
        .Cells(intCounter, 1).Value = strLine
        intCounter = intCounter + 1
    Loop
End With

objLine.Close

wksSheet.Columns("A:A").TextToColumns Destination:=wksSheet.Range("A1"), _
    DataType:=xlDelimited, semicolon:=False

End Sub
