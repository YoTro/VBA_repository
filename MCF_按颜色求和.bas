'Àà±ð=¸öÈË³£ÓÃ
'ËµÃ÷=ÎÞËµÃ÷
Option Explicit

Sub °´ÑÕÉ«ÇóºÍ()
On Error Resume Next
Dim sRng As Range
Dim cRng As Range
Dim result As Range
Dim r As Range
Dim choice As Integer

Set sRng = Selection
Dim ncolor As Long
Dim total As Double


choice = Application.InputBox("Ñ¡ÔñÍ³¼Æ·½Ê½£¬0 Îª°´±³¾°ÑÕÉ«Í³¼Æ£¬ 1Îª°´×ÖÌåÑÕÉ«Í³¼Æ", "Í³¼Æ·½Ê½", Default:=0, Type:=1)
If choice = 0 Or choice = 1 Then

Else
    MsgBox "ÎÞÐ§Ñ¡Ïî,±ØÐëÎª0 »òÕß 1"
Exit Sub

End If

Set cRng = Application.InputBox("Ñ¡ÔñÐèÒªÍ³¼ÆµÄÑÕÉ«µÄÒ»¸öµ¥Ôª¸ñ(Ö»ÐèÒ»¸öµ¥Ôª¸ñ)", "Ñ¡Ôñµ¥Ôª¸ñ", Type:=8)
If cRng Is Nothing Then Exit Sub
Set cRng = cRng.Cells(1, 1)

If choice = 0 Then  '±³¾°
    ncolor = cRng.Interior.Color
Else
    ncolor = cRng.Font.Color
End If



total = 0
For Each r In sRng
    If IsNumeric(r.Value) Then
    
    If choice = 0 Then  '±³¾°
        If r.Interior.Color = ncolor Then
         total = total + CDbl(r.Value)
        End If
    Else
        If r.Font.Color = ncolor Then
         total = total + CDbl(r.Value)
        End If
    End If
   
    End If
Next

Set result = Application.InputBox("Ñ¡Ôñ½á¹û´æ·ÅÎ»ÖÃ", "Ñ¡Ôñµ¥Ôª¸ñ", Type:=8)
If result Is Nothing Then Exit Sub
result.Value = total

End Sub
