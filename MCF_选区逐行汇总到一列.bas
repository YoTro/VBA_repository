'���=����ת��
'˵��=ѡ�����л��ܵ�һ��
Option Base 1

Sub ѡ�����л��ܵ�һ��()

    Dim arr(), count
    x = Selection.Rows.count
    y = Selection.Columns.count

    a = Selection.Value
    
    count = 0
    ReDim arr(1 To Selection.count)
    For i = 1 To x    '���Ȱ���
        For j = 1 To y
            count = count + 1
            arr(count) = a(i, j)
        Next j
    Next i
    
    Set tar = Application.InputBox(prompt:="��ѡ���Ž���ĵ�Ԫ��(��Ų��ظ�����,����)��", Title:="������", Type:=8)
    
    If tar Is Nothing Then
        Exit Sub
    End If
    
    tar.Resize(count, 1) = WorksheetFunction.Transpose(arr)  '����д��
    'tar.Resize(1, count) = WorksheetFunction.Transpose(arr)  '����д��
End Sub


