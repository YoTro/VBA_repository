'类别=
'说明=无说明
Sub 隔1列插入1列()
    Dim i, n, x
      Application.ScreenUpdating = False
       i = ActiveSheet.UsedRange.Columns.Count
        For n = i To 2 Step -1      '这里控制隔多少列插入，由步长值控制
            For x = 1 To 1            '这里控制插入多少列，由循环终止值控制
             ActiveSheet.Columns(n).Insert Shift:=xlToLeft, CopyOrigin:=xlFormatFromLeftOrAbove 'xltoright表示在右边插入
            Next x
        Next n
    Application.ScreenUpdating = True
End Sub
