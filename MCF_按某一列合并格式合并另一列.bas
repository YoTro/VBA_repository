'���=
'˵��=����B��������ݿ��ٺϲ�A�е�Ԫ��

Sub ��ĳһ�кϲ���ʽ�ϲ���һ��()
    Dim r As Range, n As Integer, beginRow As Integer
    Dim col1 As String
    Dim col2 As String
    Dim maxRow As String
    Dim lastRow As String
    Dim str
    
    str = Application.InputBox("������xx�к�yy��(�Զ��ŷֿ�)����xx�кϲ�yy��:", "����", "A,B")
    If str = False Then Exit Sub
    str = Replace(str, "��", ",")
    arr = Split(str, ",")
    
    
    col2 = arr(0)   '������
    col1 = arr(1)  'Ŀ����
    
    
    maxRow = Rows.count
    lastRow = Range(col2 & maxRow).End(xlUp).Row
    
    Range(col1 & "1:" & col1 & lastRow).MergeCells = False  'unmerge
    
    For i = 1 To lastRow
        Set r = Range(col2 & i)
        If r.MergeCells Then
            If r.MergeArea.Columns.count = 1 Then   '�ϲ����򣺵���
                If r.MergeArea.Cells.Offset.Address = r.Address Then
                    n = r.MergeArea.count
                    beginRow = r.MergeArea.Cells.Offset.Row
                    Range(col1 & beginRow & ":" & col1 & CStr(beginRow + n - 1)).Merge
                End If
            End If
        End If
    Next i
End Sub
