'类别=定位引用
'说明=返回第1行最右边非空单元的列号

Sub 返回第1行最右非空单元的列号()
X = [IV1].End(xlToLeft).Column
MsgBox X
End Sub



