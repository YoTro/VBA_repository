'类别=工作表
'说明=工作表标签排序

Sub 工作表标签排序()
Dim i As Long, j As Long, nums As Long, msg As Long

msg = MsgBox("工作表按升序排列请选 '是[Y]'. " & vbCrLf & vbCrLf & "工作表按降序排列请选 '否[N]'", vbYesNoCancel, "工作表排序")

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




