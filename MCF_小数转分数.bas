'���=��ֵת��
'˵��=��˵��
Sub С��ת����()
On Error Resume Next
Dim tar As Range
Set tar = Application.Intersect(ActiveSheet.UsedRange, Selection)
Dim r As Range
Dim tmp As String
For Each r In tar.Cells
    tmp = r.Value
    If tmp <> "" And IsNumeric(tmp) Then
        r.NumberFormatLocal = "@"  '�ı���ʽ
        r.Value = ToFenshu(tmp)
    End If
Next

End Sub

Function ToFenshu(ByVal xiaoshu As Single) As String
'С��ת��Ϊ���������С��10^-5
'�ο����£�http://club.excelhome.net/thread-810083-1-1.html
Dim i As Long, f As Single
f = xiaoshu - Int(xiaoshu)
If xiaoshu > 1 Then ToFenshu = Int(xiaoshu) & "-"
Do
i = i + 1
Loop Until Abs((i / f) - Round((i / f), 0)) < 0.1 ^ 5
ToFenshu = ToFenshu & i & "/" & Round(i / f)
End Function
