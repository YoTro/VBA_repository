'类别=
'说明=无说明
Sub 隔1行插入1行()
    Dim i, n, x
      Application.ScreenUpdating = False
       i = ActiveSheet.UsedRange.Rows.Count
        For n = i To 1 Step -1
            For x = 1 To 1
              ActiveSheet.Rows(n).Insert Shift:=xlUp, CopyOrigin:=xlFormatFromLeftOrAbove 'xldown表示在下边插入 插入的行格式随上面行的格式
            Next x
        Next n
    Application.ScreenUpdating = True
End Sub
