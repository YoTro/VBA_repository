'类别=工作簿
'说明=以当前日期为名另存文件

Sub 以当前日期为名另存文件()
ThisWorkbook.SaveAs ThisWorkbook.Path & "\" & Format(Now(), "yyyymmdd") & ".xls"
End Sub

Sub 以当前日期为名另存文件2()
ActiveWorkbook.SaveAs Filename:=Date & ".xls"
End Sub




