'类别=工作簿
'说明=以当前日期和时间为新文件名另存文件
Sub 另存文件为当前日期和时间()
    ThisWorkbook.SaveAs _ 
        ThisWorkbook.Path & "\" & Format(Now(), "yyyy" & "年" & "mm" & "月" & "dd" & "日" & "h" & "时" & "mm" & "分" & "ss" & "秒") & ".xls"
End Sub





