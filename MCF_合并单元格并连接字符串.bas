'���=�ϲ��Ͳ��
'˵��=�ϲ���Ԫ�������ַ���
Sub �ϲ���Ԫ�������ַ���()

On Error GoTo l_err
Dim Strtotal
Dim r As Range

Application.ScreenUpdating = False
Application.DisplayAlerts = False

For Each r In Selection
    Strtotal = Strtotal & r.Value
Next

Selection.Merge

With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Value = "'" & Strtotal  '�ںϲ�����ǰ�� '��
End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Exit Sub

l_err:
    MsgBox "Err: " & Err.Description

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub



