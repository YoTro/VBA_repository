'类别=个人常用
'说明=无说明
Sub 一键设置常用边框()
    With Selection.Borders
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
