'���=
'˵��=ѡ���иߺ��п�
Sub ѡ�������иߺ��п�()
    Dim str, arr
    
    str = Application.InputBox("�������иߺ��п��Զ��ŷֿ�:", "����", "10,12")
    If str = False Then Exit Sub
    
    str = Replace(str, "��", ",")
    
    arr = Split(str, ",")
    Selection.RowHeight = CInt(arr(0))   'ָ���и�
    Selection.ColumnWidth = CInt(arr(1))  'ָ���п�
    
End Sub








