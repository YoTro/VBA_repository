'���=���˳���
'˵��=����ѡ��(����Ϊ����)������ֵͬ�ĵ�Ԫ����Ŀ
Sub ������ֵͬ�ĸ���ͳ��()

Dim r As Range, tmpr As Range
Dim total As Integer

    If Selection.Columns.Count > 1 Then
        MsgBox "ѡ��ֻ�������һ���У�"
        Exit Sub
    End If

Set tmpr = Nothing

For Each r In Selection
    If tmpr Is Nothing Then
        Set tmpr = r
        total = 1
    Else
        If r.Value = tmpr.Value Then  'һ��
            total = total + 1
        Else  '��һ��
            tmpr.Offset(0, 1) = total
            Set tmpr = r
            total = 1
        End If
    End If
Next

If Not tmpr Is Nothing Then tmpr.Offset(0, 1) = total

End Sub



