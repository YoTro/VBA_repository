'���=���˳���
'˵��=��˵��
Sub Ӧ�õ׸�()
On Error Resume Next
Dim wb As Workbook
Dim ws As Worksheet
Set ws = ActiveSheet

Application.ScreenUpdating = False
Set wb = Application.Workbooks.Open("D:\�׸�\heading.xlsx")
If Not wb Is Nothing Then
    wb.Worksheets(1).Range("A1:B6").Copy ws.Range("A1")
    wb.Worksheets(1).Range("Q1:Q6").Copy ws.Range("Q1")
    wb.Worksheets(1).Range("5:6").Copy ws.Range("5:6")
    wb.Close False
    Set wb = Nothing
Else
    MsgBox "û���ҵ��ļ�"
End If
Application.ScreenUpdating = True

End Sub