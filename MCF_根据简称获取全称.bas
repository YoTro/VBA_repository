'���=���˳���
'˵��=��˵��

Sub ���ݼ�ƻ�ȡȫ��()
        
    On Error Resume Next
    Dim j1 As Range, q1 As Range
    '------------------
    Set j1 = Application.InputBox(prompt:="��ѡ������������", Title:="�����������", Type:=8)
    If j1 Is Nothing Then Exit Sub
    Set j1 = Application.Intersect(j1, j1.Worksheet.UsedRange)
    
    Set q1 = Application.InputBox(prompt:="��ѡ��ȫ����������", Title:="ȫ����������", Type:=8)
    If q1 Is Nothing Then Exit Sub
    Set q1 = Application.Intersect(q1, q1.Worksheet.UsedRange)
    
    Set rlt = Application.InputBox(prompt:="��ѡ�����������(һ����Ԫ��)", Title:="��Ž��", Type:=8)
    If rlt Is Nothing Then Exit Sub
    Set rlt = rlt.Cells(1, 1)
    
    Dim sep As String
    sep = Application.InputBox(prompt:="���ж��ƥ���������ı�����������", Title:="�������ӷ�", Default:="��", Type:=2)
    '------------------
    Dim data, jdata
    jdata = j1.Value
    data = q1.Value
    If j1.Cells.count = 1 Then 'ֻѡ��һ����Ԫ�����������⴦��
        ReDim jdata(0 To 1, 0 To 1)
        jdata(1, 1) = j1.Cells(1, 1).Value
    End If
    If q1.Cells.count = 1 Then
        MsgBox "ȫ��������ֻ��һ����Ԫ��"
        Exit Sub
    End If
    '-------------------
    Dim i, j, x1, y1
    Dim tmp As String, tmp2 As String
    For i = 1 To UBound(jdata, 1)
    For j = 1 To UBound(jdata, 2)
        tmp = jdata(i, j)
        If tmp <> "" Then
            Dim tmpr As String
            tmpr = ""
            For x1 = 1 To UBound(data, 1)
            For y1 = 1 To UBound(data, 2)
                tmp2 = data(x1, y1)
                If tmp2 <> "" Then
                    If isAllIn(tmp2, tmp) Then
                        If tmpr = "" Then
                            tmpr = tmp2
                        Else
                            tmpr = tmpr & sep & tmp2
                        End If
                    End If
                End If
            Next
            Next
            '----------------
            rlt.Offset(i - 1, j - 1).Value = tmpr
        End If
    Next
    Next
    
    MsgBox "���"
End Sub


Function isAllIn(s As String, s1 As String) As Boolean
Dim rlt As Boolean
rlt = False
Dim tmp As String
For i = 1 To Len(s1)
    If InStr(1, s, Mid(s1, i, 1)) <= 0 Then
        rlt = False
        Exit For
    Else
        rlt = True
    End If
Next

isAllIn = rlt
End Function