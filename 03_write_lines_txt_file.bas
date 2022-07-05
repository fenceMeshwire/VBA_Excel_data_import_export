Option Explicit

Sub write_lines_txt_file()

Dim strLine As String
Dim lngCounter As Long

Const output_txt = "C:\Users\...\Output.txt"

Open output_txt For Output As #1

For lngCounter = 1 To 10
  strLine = "Line: " & lngCounter
  Print #1, strLine
Next lngCounter

Close #1

End Sub
