'���=��ֵת��
'˵��=��˵��
Sub ��������������()
    Dim r As Range
    Dim str
    
    Dim bitnum As Double
    Dim tmp As Double
    '-----------------------------
    str = Application.InputBox("������Ҫ������С��λ��", "����", "2")
    If str = False Then Exit Sub
    If Not IsNumeric(str) Then Exit Sub
    
    bitnum = CDbl(str)
    If bitnum < 0 Then Exit Sub
    '-----------------------------
    For Each r In Selection
        If IsNumeric(r.Value) Then
            tmp = Application.WorksheetFunction.Round(r.Value, bitnum)
            r.Value = tmp
        End If
    Next

End Sub
