Option Explicit

Sub read_lines_txt_file()

Dim strLine As String

Const input_txt = "C:\Users\...\Input.txt"

Open input_txt For Input As #1

Do Until EOF(1)
  Line Input #1, strLine
  Debug.Print strLine
Loop

Close #1

End Sub
