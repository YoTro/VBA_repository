'类别=个人常用
'说明=无说明
Sub 区域求和并复制()

Dim objData As Object
Set objData = New DataObject
Dim I As String
I = WorksheetFunction.Sum(Selection)
objData.SetText I
objData.PutInClipboard


End Sub