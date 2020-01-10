'���=��ֵת��
'˵��=����תӢ�Ľ��

'-------------------------
 Dim StrNO(19) As String
 Dim Unit(8) As String
 Dim StrTens(9) As String
'-------------------------
Public Function ����תӢ�Ľ��()
    On Error GoTo l_err
    Dim r As Range
    Dim rlt
    
    Set r = ActiveCell
    If Not r.Value = "" Then
        rlt = NumberToString(r.Value, "��Ԫ")
        r = rlt
    End If
    Exit Function

l_err:
End Function


'-------------------------
'������������������ת��ΪӢ��
'�������� NumberToString

Public Function NumberToString(Number As Double, strMonType As String) As String
  Dim str As String, BeforePoint As String, AfterPoint As String, tmpStr As String
  Dim Point As Integer
  Dim nBit As Integer
  Dim CurString As String
  
  Call Init(strMonType)
  '"//��ʼ����
  str = CStr(Round(Number, 2))
  ' Str = Number
  If InStr(1, str, ".") = 0 Then
     BeforePoint = str
     AfterPoint = ""
  Else
     BeforePoint = Left(str, InStr(1, str, ".") - 1)
     AfterPoint = Right(str, Len(str) - InStr(1, str, "."))
     If Len(AfterPoint) = 1 Then
         AfterPoint = AfterPoint & "0" '������λС��
     End If
  End If
  
  If Len(BeforePoint) > 12 Then
     NumberToString = "Too Big."
     Exit Function
  End If
  
  str = ""
  Do While Len(BeforePoint) > 0
     nNumLen = Len(BeforePoint)
     If nNumLen Mod 3 = 0 Then
         CurString = Left(BeforePoint, 3)
         BeforePoint = Right(BeforePoint, nNumLen - 3)
     Else
         CurString = Left(BeforePoint, (nNumLen Mod 3))
         BeforePoint = Right(BeforePoint, nNumLen - (nNumLen Mod 3))
     End If
     
     nBit = Len(BeforePoint) / 3
     tmpStr = DecodeHundred(CurString)
     
     If (BeforePoint = String(Len(BeforePoint), "0") Or nBit = 0) And Len(CurString) = 3 Then
         'If CInt(Left(CurString, 1)) <> 0 And CInt(Right(CurString, 2)) <> 0 Then
             'tmpStr = Left(tmpStr, InStr(1, tmpStr, Unit(4)) + Len(Unit(4))) & Unit(8) & " " & Right(tmpStr, Len(tmpStr) - (InStr(1, tmpStr, Unit(4)) + Len(Unit(4))))
         'ElseIf CInt(Left(CurString, 1)) <> 0 And CInt(Right(CurString, 2)) = 0 Then
             'tmpStr = Unit(8) & " " & tmpStr
         'End If
     End If
     
     If nBit = 0 Then
         str = Trim(str & " " & tmpStr)
     Else
         str = Trim(str & " " & tmpStr & " " & Unit(nBit))
     End If
     
     If Left(str, 3) = Unit(8) Then str = Trim(Right(str, Len(str) - 3))
     
     If BeforePoint = String(Len(BeforePoint), "0") Then Exit Do
     Debug.Print str
  Loop
  
  BeforePoint = str
  If Len(AfterPoint) > 0 Then
     AfterPoint = Unit(6) & " " & DecodeHundred(AfterPoint) & " " & Unit(7)
  Else
     AfterPoint = Unit(5)
  End If
  NumberToString = BeforePoint & " " & AfterPoint
 End Function
 Private Function DecodeHundred(HundredString As String) As String
     Dim tmp As Integer
     If Len(HundredString) > 0 And Len(HundredString) <= 3 Then
         Select Case Len(HundredString)
             Case 1
                 tmp = CInt(HundredString)
                 If tmp <> 0 Then DecodeHundred = StrNO(tmp)
             Case 2
                 tmp = CInt(HundredString)
                 If tmp <> 0 Then
                     If (tmp < 20) Then
                         DecodeHundred = StrNO(tmp)
                     Else
                         If CInt(Right(HundredString, 1)) = 0 Then
                             DecodeHundred = StrTens(Int(tmp / 10))
                         Else
                             DecodeHundred = StrTens(Int(tmp / 10)) & " " & StrNO(CInt(Right(HundredString, 1)))  '�滻-
                         End If
                     End If
                 End If
             Case 3
                 If CInt(Left(HundredString, 1)) <> 0 Then
                     If CInt(Right(HundredString, 2)) <> 0 Then
                         DecodeHundred = StrNO(CInt(Left(HundredString, 1))) & " " & Unit(4) & " " & Unit(8) & " " & DecodeHundred(Right(HundredString, 2))
                     Else
                         DecodeHundred = StrNO(CInt(Left(HundredString, 1))) & " " & Unit(4) & " " & DecodeHundred(Right(HundredString, 2))
                     End If
                     
                 Else
                     DecodeHundred = DecodeHundred(Right(HundredString, 2))
                 End If
             Case Else
         End Select
     End If
 End Function
 Private Sub Init(ByVal strMonType As String)
  'If StrNO(1) <> "One" Then
  StrNO(1) = "One"
  StrNO(2) = "Two"
  StrNO(3) = "Three"
  StrNO(4) = "Four"
  StrNO(5) = "Five"
  StrNO(6) = "Six"
  StrNO(7) = "Seven"
  StrNO(8) = "Eight"
  StrNO(9) = "Nine"
  StrNO(10) = "Ten"
  StrNO(11) = "Eleven"
  StrNO(12) = "Twelve"
  StrNO(13) = "Thirteen"
  StrNO(14) = "Fourteen"
  StrNO(15) = "Fifteen"
  StrNO(16) = "Sixteen"
  StrNO(17) = "Seventeen"
  StrNO(18) = "Eighteen"
  StrNO(19) = "Nineteen"
  StrTens(1) = "Ten"
  StrTens(2) = "Twenty"
  StrTens(3) = "Thirty"
  StrTens(4) = "Forty"
  StrTens(5) = "Fifty"
  StrTens(6) = "Sixty"
  StrTens(7) = "Seventy"
  StrTens(8) = "Eighty"
  StrTens(9) = "Ninety"
  Unit(1) = "Thousand" '��һ����λ
  Unit(2) = "Million" '�ڶ�����λ
  Unit(3) = "Billion" '��������λ
  Unit(4) = "Hundred"
  'Unit(5) = "Only"
  'Debug.Print InStr(strMonType, "��")
  If InStr(strMonType, "��") > 0 Then
     Unit(6) = "Dollars"
     Unit(5) = "Dollars"
     Unit(7) = "Cents"
  ElseIf InStr(strMonType, "CHUSD") > 0 Then
     Unit(6) = "And Cents"
     Unit(5) = ""
     Unit(7) = "Only."
  Else
     Unit(6) = "Point"
     Unit(5) = ""
     Unit(7) = ""
  End If
  'Unit(7) = "Cents" '���ǻ��ҵĻ����Ѵ�ֵ����
  Unit(8) = "And"
  'End If
 End Sub







