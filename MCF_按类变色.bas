'类别=定位引用
'说明=区域单元分类变色
Sub 按类变色()

Dim rng As Range

For Each rng In Selection
If rng < 0 Then
rng.Interior.ColorIndex = 4   '小于0的单元变绿底色
End If
Next

For Each rng In Selection
If rng > 0 Then
rng.Interior.ColorIndex = 3    '文本、假空和大于0的单元变红底色
End If
Next

For Each rng In Selection
If rng = 0 Then
rng.Interior.ColorIndex = 2   '空值和等于0的单元变白底色
End If

Next

End Sub






