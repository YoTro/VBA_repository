'���=
'˵��=��˵��
Public Sub ������ճ����()
Dim ar, i, ii, m, str, br()
If Selection.Areas.Count > 1 Then Exit Sub
If Selection.Count = 1 Then str = Selection: GoTo 100

ar = Selection
For i = 1 To UBound(ar)
    For ii = 1 To UBound(ar, 2)
        str = Trim(str & " " & ar(i, ii))
    Next
    m = m + 1
    ReDim Preserve br(1 To m)
    br(m) = str
    str = ""
Next
str = Join(br, ",")'Join(br, Chr(10))

100:
With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
  .SetText str
  .PutInClipboard
End With

End Sub




