'���=����ͼƬ
'˵��=����ѡ�����ı������½��ı���
Sub ��ѡ���ı������½��ı���()
    For Each rag In Selection
        n = n & rag.Value & Chr(10)
    Next
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ActiveCell.Left + ActiveCell.Width, ActiveCell.Top + ActiveCell.Height, 250#, 100).Select
    Selection.Characters.Text = "���⣺" & n
    With Selection.Characters(Start:=1, Length:=3).Font
        .Name = "����"
        .FontStyle = "����"
        .Size = 12
    End With
End Sub



