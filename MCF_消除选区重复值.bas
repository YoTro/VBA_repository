'���=�ظ�ֵ�����ֵ
'˵��=��ѡ����ȡ�����ظ�������
Sub ����ѡ���ظ�ֵ()
        
    On Error Resume Next
    Dim rn As Range, res
    Dim tar
    
    Set d = CreateObject("scripting.dictionary")
    For Each rn In Selection
    If rn <> "" And Not d.exists(rn.Value) Then d.Add rn.Value, ""
    Next
    res = d.keys
    
    'For i = 0 To d.Count - 1
    	'Cells(i + 1, 5) = res(i)
    'Next
    
    Set tar = Application.InputBox(prompt:="��ѡ���Ž���ĵ�Ԫ��(��Ų��ظ�����,����)��", Title:="������", Type:=8)
    
    If tar Is Nothing Then
        Exit Sub
    End If
    
    tar.Resize(d.Count) = WorksheetFunction.Transpose(d.keys)
    'Cells(1, 11).Resize(d.Count) = WorksheetFunction.Transpose(d.keys)
End Sub

'[A:A].AdvancedFilter 2, , [e1], 1
'Selection.AdvancedFilter 2, , [s1], 1







