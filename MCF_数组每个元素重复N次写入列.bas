'���=���˳���
'˵��=����ÿ��Ԫ���ظ�N��д����
Sub ����ÿ��Ԫ���ظ�N��д����()
   
    Dim i, j, pos As Integer
    Dim strs, count, arr
    

    strs = Application.InputBox(prompt:="���������ö��Ÿ������� AA,BB,CC):", Type:=2)
    count = Application.InputBox(prompt:="����Ҫ�ظ��Ĵ���:", Type:=1)
    arr = Split(strs, ",")
    
    
    pos = 0
    For i = 0 To UBound(arr)
        For j = 1 To count
            ActiveCell.Offset(pos, 0) = arr(i)
            pos = pos + 1
        Next j
    Next i
End Sub





