'类别=打印工具
'说明=定义指定单元内容为页眉/页脚
Sub 指定单元内容为页眉页脚()
    Dim strHeader As String
    Dim strFooter As String
    strHeader = "页眉"   'Range("A1")
    strFooter = "页脚"
    
    With ActiveSheet.PageSetup
        .CenterHeader = strHeader   '定义页眉
        .CenterFooter = strFooter   '定义页脚
    End With
End Sub




