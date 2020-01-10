'类别=工作表
'说明=在目录表建立本工作簿中各表链接目录




Sub 在目录表建立各工作表链接()
Dim s%, Rng As Range
    On Error Resume Next
    Sheets("目录").Activate
    If Err = 0 Then
        Sheets("目录").UsedRange.Delete
    Else
        Sheets.Add
        ActiveSheet.Name = "目录"
    End If
    
    For i = 1 To Sheets.Count
        If Sheets(i).Name <> "目录" Then
            s = s + 1
            Set Rng = Sheets("目录").Cells(((s - 1) Mod 20) + 1, (s - 1) \ 20 + 1 + 1)
            Rng = Format(s, "  0") & ". " & Sheets(i).Name
            ActiveSheet.Hyperlinks.Add Rng, "#" & Sheets(i).Name & "!A1", ScreenTip:=Sheets(i).Name
        End If
    Next
    
    Sheets("目录").Range("b:iv").EntireColumn.ColumnWidth = 20
End Sub


