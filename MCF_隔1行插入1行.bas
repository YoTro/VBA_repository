'���=
'˵��=��˵��
Sub ��1�в���1��()
    Dim i, n, x
      Application.ScreenUpdating = False
       i = ActiveSheet.UsedRange.Rows.Count
        For n = i To 1 Step -1
            For x = 1 To 1
              ActiveSheet.Rows(n).Insert Shift:=xlUp, CopyOrigin:=xlFormatFromLeftOrAbove 'xldown��ʾ���±߲��� ������и�ʽ�������еĸ�ʽ
            Next x
        Next n
    Application.ScreenUpdating = True
End Sub
