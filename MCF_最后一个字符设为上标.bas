'类别=个人常用
'说明=无说明
Sub 最后一个字符设为上标()
    On Error Resume Next
    Dim r As Range
    If Selection.Cells.Count >= 65536 Then
        MsgBox "选择的区域太大了，超过65536个单元格"
        Exit sub
    End If
    
    Dim tar As Range
    Set tar = Selection
    Dim n As Integer
    For Each r In tar.Cells
        If r.Value <> "" Then
             n = r.Characters.Count
        
            With r.Characters(Start:=n, Length:=1).Font
                .Superscript = True
                .Subscript = False
            End With
        End If
    Next

End Sub
