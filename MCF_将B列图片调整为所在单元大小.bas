'���=����ͼƬ
'˵��=��A��������������ϵ�����B��ͼƬ��С����Ϊ���ڵ�Ԫ��С

Sub ��B��ͼƬ����Ϊ���ڵ�Ԫ��С()
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




