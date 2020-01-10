'类别=对象图片
'说明=将A列最后数据行以上的所有B列图片大小调整为所在单元大小

Sub 将B列图片调整为所在单元大小()
    Dim Pic As Picture, i&
    i = [A65536].End(xlUp).Row
    For Each Pic In Sheet1.Pictures
        If Not Application.Intersect(Pic.TopLeftCell, Range("B1:B" & i)) Is Nothing Then
            Pic.Top = Pic.TopLeftCell.Top
            Pic.Left = Pic.TopLeftCell.Left
            Pic.Height = Pic.TopLeftCell.Height
            Pic.Width = Pic.TopLeftCell.Width
        End If
    Next
End Sub




