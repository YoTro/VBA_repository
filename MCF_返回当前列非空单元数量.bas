'类别=定位引用
'说明=返回当前列非空单元数量

Sub 返回当前列非空单元数量()
    
    y = Application.CountA(Columns(ActiveCell.Column))
    MsgBox y
End Sub







