'���=�ظ�ֵ�����ֵ
'˵��=
Option Explicit


Sub ��ͬ�벻ͬ()
    On Error Resume Next
    
    Dim rn As Range
    Dim orgA, orgB
    Dim tar

    Dim dicA As Object, dicB As Object
    Dim dicA_B As Object, dicB_A As Object, dicAB As Object
    Dim ikA, ikB, i
    '--------------------------------------------------------
    Set orgA = Application.InputBox(prompt:="��ѡ��A��B�����е�A����", Title:="ѡ��A����", Type:=8)
    If orgA Is Nothing Then
        Exit Sub
    End If
    Set orgA = Intersect(ActiveSheet.UsedRange, orgA)
    
    
    Set orgB = Application.InputBox(prompt:="��ѡ��A��B�����е�B����", Title:="ѡ��B����", Type:=8)
    If orgB Is Nothing Then
        Exit Sub
    End If
    Set orgB = Intersect(ActiveSheet.UsedRange, orgB)
    '--------------------------------------------------------
    Set dicA = CreateObject("scripting.dictionary")
    Set dicB = CreateObject("scripting.dictionary")

    For Each rn In orgA
        If rn <> "" Then
            dicA.Add Trim(CStr(rn.Value)), ""
        End If
    Next
    
    For Each rn In orgB
        If rn <> "" Then
            dicB.Add Trim(CStr(rn.Value)), ""
        End If
    Next
    '--------------------------------------------------------
    Set dicAB = CreateObject("scripting.dictionary")
    Set dicA_B = CreateObject("scripting.dictionary")
    Set dicB_A = CreateObject("scripting.dictionary")
    
    ikA = dicA.keys()
    ikB = dicB.keys()
    
    
    For i = 0 To UBound(ikA)
        If dicB.exists(ikA(i)) Then
            dicAB.Add i, ikA(i) 'A��B��
        Else
            dicA_B.Add i, ikA(i) 'A��Bû��
        End If
    Next i
    
    
    For i = 0 To UBound(ikB)
        If dicA.exists(ikB(i)) Then
        Else
            dicB_A.Add i, ikB(i) 'Aû����B��
        End If
    Next i
    '--------------------------------------------------------
    Set tar = Application.InputBox(prompt:="��ѡ���Ž���ĵ�Ԫ��(����)��", Title:="������", Type:=8)
    
    If tar Is Nothing Then
        Exit Sub
    End If
    
    tar.Cells(1, 1).Offset(0, 0) = "A��B��"
    tar.Cells(1, 1).Offset(0, 1) = "A��Bû��"
    tar.Cells(1, 1).Offset(0, 2) = "B��Aû��"
    
    tar.Cells(1, 1).Offset(1, 0).Resize(dicAB.count) = WorksheetFunction.Transpose(dicAB.items)
    tar.Cells(1, 1).Offset(1, 1).Resize(dicA_B.count) = WorksheetFunction.Transpose(dicA_B.items)
    tar.Cells(1, 1).Offset(1, 2).Resize(dicB_A.count) = WorksheetFunction.Transpose(dicB_A.items)
    Exit Sub
l_err:
    MsgBox "��������" & Err.Description
    
End Sub
