'���=������
'˵��=�������ǩ����

Sub �������ǩ����()
Dim i As Long, j As Long, nums As Long, msg As Long

msg = MsgBox("����������������ѡ '��[Y]'. " & vbCrLf & vbCrLf & "����������������ѡ '��[N]'", vbYesNoCancel, "����������")

If msg = vbCancel Then Exit Sub

nums = Sheets.Count

    If msg = vbYes Then 'Sort ascending
        For i = 1 To nums
            For j = i To nums
                If UCase(Sheets(j).Name) < UCase(Sheets(i).Name) Then
                    Sheets(j).visible = true
                    Sheets(j).Move Before:=Sheets(i)
                End If
            Next j
        Next i
    Else 'Sort descending
     For i = 1 To nums
            For j = i To nums
                If UCase(Sheets(j).Name) > UCase(Sheets(i).Name) Then
                    Sheets(j).visible = true
                    Sheets(j).Move Before:=Sheets(i)
                End If
            Next j
        Next i
    End If
End Sub




