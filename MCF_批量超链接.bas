'类别=个人常用
'说明=无说明
Sub 批量超链接()
    On Error Resume Next
    Dim r1 As Range, r2 As Range, tar As Range
    Set r1 = Application.InputBox(prompt:="请选择文件路径所在区域。", Title:="文件路径", Type:=8)
    If r1 Is Nothing Then
        Exit Sub
    End If
    '----------------
    Set r2 = Application.InputBox(prompt:="请选择超链接要的显示文本所在区域。", Title:="显示文本", Type:=8)
    If r2 Is Nothing Then
        Exit Sub
    End If
    '----------------
    If r1.Rows.count = r2.Rows.count And r1.Columns.count = r2.Columns.count Then
    Else
        MsgBox "两区域的大小不一样"
        Return
    End If
    Set tar = Application.InputBox(prompt:="请选择存放结果的单元格(一个即可)。", Title:="结果存放", Type:=8)
    If tar Is Nothing Then
        Exit Sub
    End If
    tar = tar.Resize(r1.Rows.count, r1.Columns.count)
    '----------------
    Dim i, j
    Dim txt1 As String, txt2 As String
    For i = 1 To r1.Rows.count
        For j = 1 To r1.Columns.count
            txt1 = r1.Cells(i, j).Value
            txt2 = r2.Cells(i, j).Value
            ActiveSheet.Hyperlinks.Add Anchor:=tar.Cells(i, j), Address:=txt1, TextToDisplay:=txt2
        Next
    Next
    '----------------
  
   MsgBox "完成"
End Sub