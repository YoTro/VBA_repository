'���=��ֵ��ֵ
'˵��=��˵��
Sub ѡ���жϿ���()
    On Error Resume Next
    Dim arr, flag, rlt
    Dim i, j
    Dim str, chc As Integer
    
    Dim tar As Range, DataRng As Range
    Dim colName, beginRow As Long
    
    '-------------------------------
    Set DataRng = Intersect(Selection, Selection.Worksheet.UsedRange)
    
    If Selection.Cells.count <= 1 Then
        MsgBox "��ѡ���һ�������"
        Exit Sub
    End If

    If DataRng.Areas.count > 1 Then
        MsgBox "ѡ��ֻ����һ������"
        Exit Sub
    End If

    beginRow = DataRng.Cells(1, 1).Row
    '-------------------------------
    str = Application.InputBox("�жϿ��б�׼: " & vbCrLf & vbCrLf & _
                                "1.��ѡ��ÿһ����ֻҪ��һ����Ԫ��Ϊ��    " & vbCrLf & vbCrLf & _
                                        "2.��ѡ��ÿһ����ȫ����Ԫ��Ϊ��", "��ѡ��", "1")
    If str = False Then Exit Sub
    If Not IsNumeric(str) Then Exit Sub
    
    chc = CInt(str)
    If chc <> 1 And chc <> 2 Then Exit Sub
    '-------------------------------
    arr = DataRng
    ReDim rlt(1 To UBound(arr))
    
    For i = 1 To UBound(arr)
        flag = 0
        rlt(i) = ""
        '----------------
        For j = 1 To UBound(arr, 2)
            If Trim(arr(i, j)) = "" Then flag = flag + 1
        Next j
        '----------------
        If chc = 1 Then
            If flag >= 1 Then rlt(i) = "����"
        ElseIf chc = 2 Then
            If flag >= UBound(arr, 2) Then rlt(i) = "����"
        End If
    Next i
    '-------------------------------
    Set tar = Application.InputBox(prompt:="��ѡ��һ�հ������ڴ�Ž����", Title:="������", Type:=8)
    
    If tar Is Nothing Then
        Exit Sub
    End If
    '---------------------
    colName = tar.Column
    '---------------------
    For i = 1 To UBound(rlt)
        Cells(beginRow + i - 1, colName) = rlt(i)
    Next
End Sub



