'类别=打印工具
'说明=返回总页码
Sub 返回总页码()
    Dim a
    'Sheet1.Activate
    a = ExecuteExcel4Macro("Get.Document(50)")
    Msgbox "总页码:" & a
End Sub





