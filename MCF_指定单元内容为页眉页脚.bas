'���=��ӡ����
'˵��=����ָ����Ԫ����Ϊҳü/ҳ��
Sub ָ����Ԫ����Ϊҳüҳ��()
    Dim strHeader As String
    Dim strFooter As String
    strHeader = "ҳü"   'Range("A1")
    strFooter = "ҳ��"
    
    With ActiveSheet.PageSetup
        .CenterHeader = strHeader   '����ҳü
        .CenterFooter = strFooter   '����ҳ��
    End With
End Sub




