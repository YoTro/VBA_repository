'���=�ظ�ֵ�����ֵ
'˵��=��ѡ���е��ظ�ֵȾɫ

Sub ���ѡ���ظ�ֵ()
    On Error Resume Next
    Dim rn As Range, first As Range
    Dim ColorIdx As Integer
    
    Set d = CreateObject("scripting.dictionary")
    Selection.Interior.ColorIndex = 2
    
    ColorIdx = 0
    For Each rn In Selection
        If rn <> "" Then
            If d.exists(rn.Value) Then
                Set first = Range(d(rn.Value))  '��һ�γ��ֵĵ�Ԫ��
                If first.Interior.ColorIndex = 2 Then  '��һ�γ���ʱ δ���ù���ɫ
                    '----------------------------------
                    ColorIdx = (ColorIdx + 1) Mod 56 + 1  '��ɫ��ѡ��Χ��0~56
                    If ColorIdx = 2 Then ColorIdx = 3
                    '----------------------------------
                    first.Interior.ColorIndex = ColorIdx
                Else
                    ColorIdx = first.Interior.ColorIndex
                End If
                rn.Interior.ColorIndex = ColorIdx
            Else
                d.Add rn.Value, rn.Address
            End If
        End If
    Next

End Sub




