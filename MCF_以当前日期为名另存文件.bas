'���=������
'˵��=�Ե�ǰ����Ϊ������ļ�

Sub �Ե�ǰ����Ϊ������ļ�()
ThisWorkbook.SaveAs ThisWorkbook.Path & "\" & Format(Now(), "yyyymmdd") & ".xls"
End Sub

Sub �Ե�ǰ����Ϊ������ļ�2()
ActiveWorkbook.SaveAs Filename:=Date & ".xls"
End Sub




