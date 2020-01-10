Sub CommentPic()
    '根据名称一键将图片批量插入到单元格的批注中去
    '1，代码运行后，首先弹出操作界面，要求用户选择图片所存放的文件夹，注意是选择文件夹，不是选择图片，双击打开文件夹后看不到图片是正常现象……。

    '2，图片的名称需要和单元格的值相匹配；小代码如果找不到单元格的值所对应的图片，会直接跳过，处理下一个单元格；代码运行结束后，会发出信息，告知一共处理成功了几个图片以及未成功几个非空单元格的图片。
    '3，代码中使用了intersect语句，Set Rng = Intersect(Rng.Parent.UsedRange, Rng)，因此用户可以选择整列（比如整个A列）或多列单元格区域运行代码，而不用担心因为无谓的运算量过大，造成程序假死的情况。

    '4，代码导入的图片格式支持五种常见的类型，Arr = Array(".jpg", ".jpeg", ".bmp", ".png", ".gif")

    '5，批注图片的高度和宽度可以根据自身情况做调整，相关代码如下：

     '.Shape.Height = 150 \'图形的高度，可以根据需要自己调整

    '.Shape.Width = 150 \'图形的宽度，可以根据需要自己调整
    Dim Arr, i&, k&, n&, pd&

    Dim PicName$, PicPath$, FdPath$

    Dim Rng As Range, Cll As Range

    On Error Resume Next

    '用户选择图片所在的文件夹

    With Application.FileDialog(msoFileDialogFolderPicker)

        .AllowMultiSelect = False '不允许多选

       If .Show Then FdPath = .SelectedItems(1) Else: Exit Sub

    End With

    If Right(FdPath, 1) <> "\" Then FdPath = FdPath & "\"

    Set Rng = Application.InputBox("请选择需要插入图片到批注中的单元格区域", Type:=8)

    '用户选择需要插入图片到批注中的单元格或区域

    If Rng.Count = 0 Then Exit Sub

    Set Rng = Intersect(Rng.Parent.UsedRange, Rng)

    'intersect语句避免用户选择整列单元格，造成无谓运算的情况

    Arr = Array(".jpg", ".jpeg", ".bmp", ".png", ".gif")

    '用数组变量记录五种文件格式

    Application.ScreenUpdating = False

    For Each Cll In Rng

    '遍历选择区域的每一个单元格

        Cll.Comment.Delete '删除旧的批注

        PicName = Cll.Text '图片名称

        If Len(PicName) Then '如果单元格存在值

            PicPath = FdPath & PicName '图片路径

            pd = 0 'pd变量标记是否找到相关图片

            For i = 0 To UBound(Arr)

            '由于不确定用户的图片格式，因此遍历图片格式

                If Len(Dir(PicPath & Arr(i))) Then

                '如果存在相关文件

                    Cll.AddComment '增加批注

                    With Cll.Comment

                        .Visible = True '批注可见

                        .Text Text:=""

                        .Shape.Select True '选中批注图形

                        Selection.ShapeRange.Fill.UserPicture PicPath & Arr(i)

                        '插入图片到批注中

                        .Shape.Height = 150 '图形的高度，可以根据需要自己调整

                        .Shape.Width = 150 '图形的宽度，可以根据需要自己调整

                        .Visible = False '取消显示

                    End With

                    pd = 1 '标记找到结果

                    n = n + 1 '累加找到结果的个数

                    Exit For '找到结果后就可以退出文件格式循环

                End If

            Next

            If pd = 0 Then k = k + 1 '如果没找到图片累加个数

        End If

    Next

    MsgBox "共处理成功" & n & "个图片，另有" & k & "个非空单元格未找到对应的图片。"

    Application.ScreenUpdating = True

End Sub

