Sub Merge()
' 合并所有工作簿
' Merge Macro
'

Dim FileOpen
Dim X As Integer
Application.ScreenUpdating = False
FileOpen = Application.GetOpenFilename(FileFilter:="Microsoft Excel文件(.xlsx),.xlsx", MultiSelect:=True, Title:="合并工作薄")
X = 1
While X <= UBound(FileOpen)
Workbooks.Open Filename:=FileOpen(X)
Sheets().Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
X = X + 1
Wend
ExitHandler:
Application.ScreenUpdating = True
Exit Sub
errhadler:
MsgBox Err.Description
End Sub
