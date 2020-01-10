'类别=工作表
'说明=合并各工作表内容

Sub 合并各工作表内容()
sp = InputBox("各表内容之间，间隔几行？不输则默认为0")
If sp = "" Then
  sp = 0
End If

st = InputBox("各表从第几行开始合并？不输则默认为2")
If st = "" Then
   st = 2
End If

Sheets(1).Select
Sheets.Add
  
  If st > 1 Then
    Sheets(2).Select
    Rows("1:" & CStr(st - 1)).Select
    Selection.Copy
    Sheets(1).Select
    Range("A1").Select
    ActiveSheet.Paste
  y = st - 1
  End If
  
For i = 2 To Sheets.Count
    Sheets(i).visible = true
  Sheets(i).Select
     For v = 1 To 256
        zd = Cells(65535, v).End(xlUp).Row
        If zd > x Then
           x = zd
        End If
     Next v

  If y + x - st + 1 + sp > 65536 Then
  MsgBox "内容太多，仅合并前" & i - 2 & "个表的内容，请把其它表复制到新工作薄里再用此程序合并！"
  Else:
  
  Rows(st & ":" & x).Select
  Selection.Copy
  Sheets(1).Select
  Range("A" & CStr(y + 1)).Select
  ActiveSheet.Paste
  
  Sheets(i).Select
  Range("A1").Select                        '取消单元格被全选状态。
  Application.CutCopyMode = False           '忘掉复制的内容。
  End If
  
  y = y + x - st + 1 + sp
  x = 0
Next i

Sheets(1).Select
Range("A1").Select                          '光标移至A1。
MsgBox "这就是合并后的表，请命名！"

End Sub







