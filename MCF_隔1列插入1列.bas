'���=
'˵��=��˵��
Sub ��1�в���1��()
    Dim i, n, x
      Application.ScreenUpdating = False
       i = ActiveSheet.UsedRange.Columns.Count
        For n = i To 2 Step -1      '������Ƹ������в��룬�ɲ���ֵ����
            For x = 1 To 1            '������Ʋ�������У���ѭ����ֵֹ����
             ActiveSheet.Columns(n).Insert Shift:=xlToLeft, CopyOrigin:=xlFormatFromLeftOrAbove 'xltoright��ʾ���ұ߲���
            Next x
        Next n
    Application.ScreenUpdating = True
End Sub
