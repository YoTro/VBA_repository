'���=�ظ�ֵ�����ֵ
'˵��=���ظ������������
Sub ���ɿ��ظ��������()
    On Error Resume Next
    
    Dim count As Long
    Dim rn As Range
    Dim max, min, unit As Double
    Dim bRepeat As Boolean
    Dim d

    count = Selection.Cells.count   '16384��
    If count >= Columns.count Then
        MsgBox "�벻Ҫѡ��̫�������"
        Exit Sub
    End If

    min = Application.InputBox(prompt:="�����С����", Type:=1, Default:="0")
    If Not IsNumeric(min) Then Exit Sub    'false  ����׼ȷ��0Ҳ��false
    
    max = Application.InputBox(prompt:="����������", Type:=1, Default:="100")
    If Not IsNumeric(max) Then Exit Sub
    
    unit = Application.InputBox(prompt:="������ľ�ȷ��λ,�羫ȷ��1����ȷ��0.2 �ȵ�", Type:=1, Default:="1")
    If Not IsNumeric(unit) Then Exit Sub  'If unit = False Then Exit Sub
    
    Randomize Timer
    For Each rn In Selection
        rn = Int(Rnd() * Int((max - min + unit) / unit)) * unit + min
    Next

End Sub


