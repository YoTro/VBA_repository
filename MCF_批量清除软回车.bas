'类别=批量删除
'说明=批量清除软回车

Sub 批量清除软回车()
      '也可直接使用Alt+10或13替换
    Cells.Replace What:=Chr(10), Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub





