'���=��λ����
'˵��=ѡ����ѡ

Sub ѡ����ѡ()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Dim raddress As String, taddress As String
raddress = Selection.Address
taddress = ActiveSheet.UsedRange.Address
With Sheets.Add
.Range(taddress) = 0
.Range(raddress) = "=0"
raddress = .Range(taddress).SpecialCells(xlCellTypeConstants, 1).Address
.Delete
End With
ActiveSheet.Range(raddress).Select
Application.ScreenUpdating = True
End Sub





