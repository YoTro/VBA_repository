'类别=工作表
'说明=以指定区域为表目录补充新表

Sub 以指定区域为表目录补充新表()
    Dim dic As Object, sh As Worksheet
    Dim arr, item
    arr = Range("B1:BB1")
    Set dic = CreateObject("scripting.dictionary")
    For Each sh In ThisWorkbook.Worksheets
        dic.Add sh.Name, ""
    Next
    For Each item In arr
        If item <> "" And Not dic.exists(Trim(item)) Then
            With ThisWorkbook.Worksheets.Add
                 .Name = item
            End With
        End If
    Next
    Set dic = Nothing
End Sub


