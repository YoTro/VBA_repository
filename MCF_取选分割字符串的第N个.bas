'���=���˳���
'˵��=
Option Explicit


Sub ȡѡ�ָ��ַ����ĵ�N��()
        
    On Error Resume Next
    Dim arr, tmp, rlt
    Dim x, count
    Dim pos As Integer
    Dim sepor, str
    Dim i
    Dim tar As Range
    
    If Selection.Columns.count > 1 Then Exit Sub
    '----------------------------------------------------
    sepor = Application.InputBox("��������ʲô��Ϊ�ָ��ַ���:", "����", " ")
    If sepor = False Then Exit Sub
    
    str = Application.InputBox("��������ȡ�ڼ���:", "����", "1")
    If str = False Then Exit Sub
    If Not IsNumeric(str) Then Exit Sub
    
    
    pos = CInt(str) - 1
    arr = Selection
    ReDim rlt(1 To UBound(arr))
    
    For i = 1 To UBound(arr)
        tmp = arr(i, 1)
        tmp = regReplace(tmp, "(" & sepor & ")+", sepor)
        
        x = Split(tmp, sepor)
        
        If UBound(x) >= pos Then
            rlt(i) = x(pos)
        Else
            rlt(i) = arr(i, 1)
        End If
        
    Next i

    '------------------------------------------------
    Set tar = Application.InputBox(prompt:="��ѡ���Ž���ĵ�Ԫ��(����)��", Title:="������", Type:=8)
    
    If tar Is Nothing Then
        Exit Sub
    End If
    
    tar.Resize(UBound(rlt)) = WorksheetFunction.Transpose(rlt)
End Sub





Function regReplace(ByVal v As String, ByVal reg As String, ByVal rp As String) As String
    Dim rlt As String
    
    Dim re As Object
    Dim mhs
    Dim mh
    
    Set re = CreateObject("VBScript.regexp") ' regexp  '����new
    re.Pattern = reg
    're.IgnoreCase = True
    re.Global = True  '���ƥ��ʱ��Ҫ
    
    rlt = re.Replace(v, rp)  '�滻
    
    regReplace = rlt
End Function

