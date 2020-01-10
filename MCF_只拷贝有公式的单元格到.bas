'类别=个人常用
'说明=无说明
Sub 只拷贝有公式的单元格到()
        
    On Error Resume Next
    Dim rn As Range
    Dim i, j As Long
    Dim tar, m As Range
    Set m = Application.Intersect(Application.ActiveSheet.UsedRange, Selection)
    If m Is Nothing Then Exit Sub
    If m.Cells.count <= 1 Then Exit Sub
    
    Set tar = Application.InputBox(prompt:="请选择要粘贴的起始区域:(选一个单元格)", Title:="结果存放", Type:=8)
    If tar Is Nothing Then Exit Sub
    Set tar = tar.Cells(1, 1)
    
    For i = 1 To m.Rows.count
        For j = 1 To m.Columns.count
            If m.Cells(i, j).HasFormula Then
                m.Cells(i, j).Copy tar.Offset(i - 1, j - 1)
            End If
        Next
    Next
    
    MsgBox "完成"
End Sub