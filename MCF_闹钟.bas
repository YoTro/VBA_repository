'类别=
'说明=闹钟。到指定时间执行宏，或延迟几秒执行。需要点击 添加到工作簿后执行

Sub 闹钟()
    Application.OnTime ("11:45:00"), "提示1"    '宏名字
    Application.OnTime Now + TimeValue("00:00:15"), "提示2"
End Sub

Sub 提示1()
    msgbox "提示，闹钟"
End Sub

Sub 提示2()
    msgbox "提示，闹钟"
End Sub





