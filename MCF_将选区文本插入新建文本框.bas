'类别=对象图片
'说明=将所选区域文本插入新建文本框
Sub 将选区文本插入新建文本框()
    For Each rag In Selection
        n = n & rag.Value & Chr(10)
    Next
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ActiveCell.Left + ActiveCell.Width, ActiveCell.Top + ActiveCell.Height, 250#, 100).Select
    Selection.Characters.Text = "问题：" & n
    With Selection.Characters(Start:=1, Length:=3).Font
        .Name = "黑体"
        .FontStyle = "常规"
        .Size = 12
    End With
End Sub



