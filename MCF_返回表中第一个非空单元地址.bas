'类别=定位引用
'说明=返回表中第一个非空单元地址(行搜索)




Sub 返回表中第一个非空单元地址()
MsgBox Cells.Find("*").Address
End Sub


