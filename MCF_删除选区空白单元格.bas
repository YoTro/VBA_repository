'���=��ֵ��ֵ
'˵��=��˵��
Sub ɾ��ѡ���հ׵�Ԫ��()
    On Error Resume Next
    Dim r As Range, tmp As Range
    Set r = Intersect(ActiveSheet.UsedRange, Selection)
    
    Dim dir As Long
    dir = Application.InputBox("ɾ����   0�������ƶ�     1�������ƶ���", Default:=0, Type:=1)
    If dir = 0 Then
        dir = xlShiftUp
    Else
        dir = xlShiftToLeft
    End If
    
    Application.ScreenUpdating = False
    For i = r.Cells.Rows.Count To 1 Step -1
        For j = r.Cells.Columns.Count To 1 Step -1
            Set tmp = r.Cells(i, j)
            If tmp.Value = "" Then
                 tmp.Delete dir  'xlShiftToLeft   xlShiftUp
            End If
        Next j
    Next i
    
    Application.ScreenUpdating = True
End Sub