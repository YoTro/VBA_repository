Sub SplitShts()
'一键根据筛选条件将总表数据拆分为多个分表
    Dim d As Object, sht As Worksheet
    Dim aData, aResult, aTemp, aKeys, i&, j&, k&, x&
    Dim rngData As Range, rngGist As Range
    Dim lngTitleCount&, lngGistCol&, lngColCount&
    Dim rngFormat As Range
    Dim strKey As String
    Set d = CreateObject("scripting.dictionary")
    Set rngGist = Application.InputBox("请框选拆分依据列！只能选择单列单元格区域！", Title:="提示", Type:=8)
    '========用户选择的拆分依据列
    lngGistCol = rngGist.Column
    '========拆分依据列的列标
    lngTitleCount = Val(Application.InputBox("请输入总表标题行的行数？"))
    '========用户设置总表的标题行数
    If lngTitleCount < 0 Then MsgBox "标题行数不能为负数，程序退出。": Exit Sub
    Set rngData = ActiveSheet.UsedRange
    '========总表的数据区域
    Set rngFormat = ActiveSheet.Cells
    '========总表的单元格集用于粘贴总表格式
    aData = rngData.Value
    lngGistCol = lngGistCol - rngData.Column + 1
    '========计算依据列在数组中的位置
    lngColCount = UBound(aData, 2)
    '========数据源的列数
    For i = lngTitleCount + 1 To UBound(aData)
        If aData(i, lngGistCol) = "" Then aData(i, lngGistCol) = "单元格空白"
        strKey = aData(i, lngGistCol)
    '========统一转换为字符串格式
        If Not d.exists(strKey) Then
    '========字典中不存在关键字时将行号装入字典
            d(strKey) = i
        Else
            d(strKey) = d(strKey) & "," & i
    '========如果字段存在关键字则合并行号
        End If
    Next
    Application.DisplayAlerts = False
    For Each sht In ActiveWorkbook.Worksheets
    '========删除字典中存在的表名
        If d.exists(sht.Name) Then sht.Delete
    Next
    Application.DisplayAlerts = True
    aKeys = d.keys
    '========字典的key集
    Application.ScreenUpdating = False
    For i = 0 To UBound(aKeys)
        If aKeys(i) <> "" Then
            aTemp = Split(d(aKeys(i)), ",")
    '========取出item里储存的行号
            ReDim aResult(1 To UBound(aTemp) + 1, 1 To lngColCount)
    '========声明放置结果的数组aResult
            k = 0
            For x = 0 To UBound(aTemp)
                k = k + 1
                For j = 1 To lngColCount
                    aResult(k, j) = aData(aTemp(x), j)
                Next
            Next
            With Worksheets.Add(, Sheets(Sheets.Count))
    '========新建一个工作表
                .Name = aKeys(i)
                .[a1].Resize(UBound(aData), lngColCount).NumberFormat = "@"
    '========设置单元格为文本格式
                If lngTitleCount > 0 Then .[a1].Resize(lngTitleCount, lngColCount) = aData
    '========标题行
                .[a1].Offset(lngTitleCount, 0).Resize(k, lngColCount) = aResult
    '========数据
                rngFormat.Copy
                .[a1].PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    '========复制粘贴总表的格式
                .[a1].Offset(lngTitleCount + k, 0).Resize(UBound(aData) - k - lngTitleCount, 1).EntireRow.Delete
    '========删除多余的格式单元格
                .[a1].Select
            End With
        End If
    Next
    rngData.Parent.Activate
    '========激活总表
    Application.ScreenUpdating = True
    Set d = Nothing
    Set rngData = Nothing
    Set rngGist = Nothing
    Set rngFormat = Nothing
    Erase aData: Erase aResult
    MsgBox "数据拆分完成！"
End Sub