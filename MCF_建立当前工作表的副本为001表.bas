'类别=工作表

'说明=建立当前工作表的副本为001表



Sub 建立当前工作表的副本为001表()
    ActiveSheet.Copy Before:=Sheets(1)
    ActiveSheet.Name = "001"
End Sub

