'���=�ظ�ֵ�����ֵ
'˵��=���ɲ��ظ��������

Sub ���ɲ��ظ��������()
    On Error Resume Next
    
    Dim count As Long, needCount As Long
    Dim rn As Range
    Dim max, min, unit As Double
    Dim bRepeat As Boolean
    Dim d As Object
    Dim i, v

    count = Selection.Cells.count
    If count > 10000 Then
        MsgBox "�벻Ҫѡ�񳬹�10000����Ԫ��"
        Exit Sub
    End If

    min = Application.InputBox(prompt:="�����С����", Type:=1, Default:="0")
    If Not IsNumeric(min) Then Exit Sub    'false  ����׼ȷ��0Ҳ��false
    
    max = Application.InputBox(prompt:="����������", Type:=1, Default:="100")
    If Not IsNumeric(max) Then Exit Sub
    
    unit = Application.InputBox(prompt:="������ľ�ȷ��λ,�羫ȷ��1����ȷ��0.2 �ȵ�", Type:=1, Default:="1")
    If Not IsNumeric(unit) Then Exit Sub  'If unit = False Then Exit Sub
    
    '---------------------------------------
    needCount = Int((max - min + unit) / unit)
    If count > needCount Then
        count = needCount  '�������
        'MsgBox "��ѡ�������̫���޷����ɲ��ظ���������� ����ֻ��ѡ��" & needCount & "����Ԫ��"
        'Exit Sub
    End If
    
    Set d = CreateObject("scripting.dictionary")
    Randomize Timer
    For i = 1 To count
        Do
            v = Int(Rnd() * Int((max - min + unit) / unit)) * unit + min
        Loop While d.exists(v)

        Selection.Cells(i) = v
        d.Add v, ""
    Next i
End Sub





