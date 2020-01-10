Sub Merge_sheet()()


    Dim sht As Worksheet, rng As Range, k&, trow&

    Application.ScreenUpdating = False

    '取消屏幕更新，加快代码运行速度

    trow = Val(InputBox("请输入标题的行数", "提醒"))

    If trow < 0 Then MsgBox "标题行数不能为负数。", 64, "警告": Exit Sub

    '取得用户输入的标题行数，如果为负数，退出程序

    Cells.ClearContents

    '清空当前表数据

    Cells.NumberFormat = "@"

    '设置文本格式

    For Each sht In Worksheets

    '遍历表格

        If sht.Name <> ActiveSheet.Name Then

        '如果表格名称不等于当前表名则进行复制数据……

            Set rng = sht.UsedRange

            '定义rng为表格已用区域

            k = k + 1

            '累计K值

            If k = 1 Then

            '如果是首个表格，则K为1，则把标题行一起复制到汇总表

                rng.Copy

                [a1].PasteSpecial Paste:=xlPasteValues

            Else

                '否则，扣除标题行后再复制黏贴到总表，只黏贴数值

                rng.Offset(trow).Copy

                Cells(ActiveSheet.UsedRange.Rows.Count + 1, 1).PasteSpecial Paste:=xlPasteValues

            End If

        End If

    Next

    [a1].Activate

    '激活A1单元格

    Application.ScreenUpdating = True

    '恢复屏幕刷新

End Sub