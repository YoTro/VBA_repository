'类别=打印工具
'说明=奇偶页分别打印

Sub 奇偶页分别打印()
  Dim i%, Ps%
  Ps = ExecuteExcel4Macro("GET.DOCUMENT(50)") '总页数
  MsgBox "现在打印奇数页,按确定开始."
  For i = 1 To Ps Step 2
    ActiveSheet.PrintOut from:=i, To:=i
  Next i
  MsgBox "现在打印偶数页,按确定开始."
  For i = 2 To Ps Step 2
    ActiveSheet.PrintOut from:=i, To:=i
  Next i
End Sub




