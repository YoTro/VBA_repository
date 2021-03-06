'类别=个人常用
'说明=无说明
Sub 拆分本表() '逐行复制，速度偏慢，通用性好
Dim SplitCol As String, ColNum As Integer, HeadRows As Byte
Dim arr, lastrow, i, ShtIndex
Dim only
Set only = CreateObject("scripting.dictionary") 'Set only = New Collection
'-------------
'指定拆分条件所在列。可以根据实际情况修改列标
Dim tmpX
tmpX = Application.InputBox("请输入拆分条件所在列:", "指定拆分条件所在列", "E", Type:=2)
If tmpX = False Then Exit Sub
SplitCol = tmpX

'指定标题行数，该区域不参与拆分
tmpX = Application.InputBox("指定标题行数，该区域不参与拆分", "标题行数", "1", Type:=1)
If tmpX = False Then Exit Sub
HeadRows = tmpX
'-----------------
If HeadRows >= ActiveSheet.UsedRange.Rows.Count Then Exit Sub '如果指定的标题行大于已用区域行数则退出程序
ColNum = Cells(1, SplitCol).Column  '将列标转换成数字
lastrow = ActiveSheet.UsedRange.Rows.Count  '获取当前表已用区域的行数
arr = Range(Cells(HeadRows + 1, SplitCol), Cells(lastrow, SplitCol)).Value  '将拆分列的数据赋与变量arr
'-----------------
On Error Resume Next
For i = 1 To lastrow - HeadRows  '遍历arr所有数据
  '提取其中的不重复值
  If Len(arr(i, 1)) > 0 Then only.Add CStr(arr(i, 1)), CStr(arr(i, 1))
Next i
ShtIndex = ActiveSheet.Index  '获取当前表位置
'-----------------
Dim ikeys
ikeys = only.keys
'-----------------
On Error Resume Next
For i = 0 To only.Count - 1
    Debug.Print Sheets(ikeys(i)).Name  '获取与only对象中每个元素同名的工作表名（用意为判断是否存在该工作表）
    If Err = 0 Then MsgBox "当前工作簿已存在与待拆分项目同名的工作表""" & ikeys(i) & """，暂无法拆分", 64, "友情提示": Exit Sub
    Err.Clear
Next i
'-----------------
Application.ScreenUpdating = False  '关闭屏幕更新，加快执行速度
Application.Calculation = xlCalculationManual  '调为手动计算，加快执行速度
For i = 0 To only.Count - 1 '创建工作表，表的数量与表名由only对象中不重复值而定
    Sheets.Add After:=Sheets(Sheets.Count)  '创建
    Sheets(Sheets.Count).Name = ikeys(i)    '命名
    Sheets(ShtIndex).Rows("1:" & HeadRows).Copy Sheets(Sheets.Count).Cells(1, 1)  '复制标题
Next i
'-----------------
Sheets(ShtIndex).Select  '返回被拆分的工作表
For i = HeadRows + 1 To lastrow         '逐行复制数据
  If Len(Cells(i, SplitCol)) > 0 Then  '排除空值
    With Sheets(Cells(i, SplitCol).Text).UsedRange.Rows(Sheets(Cells(i, SplitCol).Text).UsedRange.Rows.Count + 1)
          Rows(i).Copy .Cells(1)  '第一次复制，复制所有数据，仅取其格式
          .Cells = Rows(i & ":" & i).Value  '第二次复制，仅复制数值
    End With
  End If
Next i   '第一列为空时，会有bug
'-----------------
Application.ScreenUpdating = True  '恢复屏幕更新
Application.Calculation = xlCalculationAutomatic  '恢复自动计算
MsgBox "拆分完毕！", 64, "友情提示"
End Sub
