'���=������
'˵��=��Ŀ¼�������������и�������Ŀ¼




Sub ��Ŀ¼����������������()
Dim s%, Rng As Range
    On Error Resume Next
    Sheets("Ŀ¼").Activate
    If Err = 0 Then
        Sheets("Ŀ¼").UsedRange.Delete
    Else
        Sheets.Add
        ActiveSheet.Name = "Ŀ¼"
    End If
    
    For i = 1 To Sheets.Count
        If Sheets(i).Name <> "Ŀ¼" Then
            s = s + 1
            Set Rng = Sheets("Ŀ¼").Cells(((s - 1) Mod 20) + 1, (s - 1) \ 20 + 1 + 1)
            Rng = Format(s, "  0") & ". " & Sheets(i).Name
            ActiveSheet.Hyperlinks.Add Rng, "#" & Sheets(i).Name & "!A1", ScreenTip:=Sheets(i).Name
        End If
    Next
    
    Sheets("Ŀ¼").Range("b:iv").EntireColumn.ColumnWidth = 20
End Sub


