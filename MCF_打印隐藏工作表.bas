'类别=打印工具
'说明=对指定工作表执行取消隐藏》打印》隐藏工作表
Sub 打印隐藏工作表()
    On Error GoTo l_err
    Dim wsName As String
    wsName = "报表1"  '工作表名
    
    Sheets(wsName).Visible = True
    Sheets(wsName).PrintOut Copies:=1, Collate:=True
    Sheets(wsName).Visible = False
    Exit Sub
l_err:
    msgbox "发生错误：" & Err.Description
End Sub






