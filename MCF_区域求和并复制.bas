'���=���˳���
'˵��=��˵��
Sub ������Ͳ�����()

Dim objData As Object
Set objData = New DataObject
Dim I As String
I = WorksheetFunction.Sum(Selection)
objData.SetText I
objData.PutInClipboard


End Sub