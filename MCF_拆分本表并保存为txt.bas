'类别=个人常用
'说明=无说明
Option Explicit

Sub 拆分本表并保存为txt()
On Error Resume Next
Dim ws As Worksheet, tar As Range, r1 As Range
Set ws = ActiveSheet
Set tar = ws.UsedRange
'分割符， 保存路径
'用空行代表一个新的文件
Dim folder As String
folder = PickFolder()
If folder = "" Then Exit Sub
'--------------
Dim sepor As String
Dim tmp As String
tmp = Application.InputBox("用什么作为分割符,请选择数字选项：  0:空格  1:tab  2:逗号  ", "分割符", "0", Type:=1)
If tmp = "" Then Exit Sub
If tmp = "0" Then
    sepor = " "
ElseIf tmp = "1" Then
    sepor = vbTab
ElseIf tmp = "2" Then
    sepor = ","
Else
    sepor = tmp
End If
'--------------
Dim iStart As Long, iEnd As Long
iStart = 1
Do
    iStart = GetStartDataRow(tar, iStart)
    If iStart <= 0 Then Exit Do
    iEnd = GetEndDataRow(tar, iStart)
    If iEnd <= 0 Then Exit Do
    'MsgBox iStart & " --  " & iEnd
    '--------------
    Set r1 = tar.Worksheet.Range(tar.rows(iStart), tar.rows(iEnd))
    DealOneArea r1, folder, sepor
    '--------------
    iStart = iEnd + 1
    If iEnd >= tar.rows.Count Then Exit Do
Loop While 1


End Sub

Private Sub DealOneArea(tar As Range, folder As String, sepor As String)
    On Error Resume Next
    Dim fn As String, fp As String
    Dim i, j As Long
    fn = tar.Cells(1, 1).Value
    fp = folder & "\" & Trim(fn) & ".txt"
    Dim data
    data = tar.Value
    Dim rows As Long, cols As Long
    rows = tar.rows.Count: cols = tar.Columns.Count
    Dim rlt As String, tmp As String
    For i = 1 To rows
        tmp = ""
        For j = 2 To cols
            tmp = IIf(j = 2, data(i, j), tmp & sepor & data(i, j))
        Next
        rlt = IIf(i = 1, tmp, rlt & vbCrLf & tmp)
    Next
    '---------------
    'MsgBox Len(rlt)
    WriteTxt fp, rlt
End Sub

Private Function GetStartDataRow(tar As Excel.Range, iRow As Long) As Long
Dim rlt As Long '有数据的第一行，就是起始行
rlt = 0
Dim i As Long
For i = iRow To tar.rows.Count
    If Application.WorksheetFunction.CountA(tar.rows(i)) > 0 Then
        rlt = i
        Exit For
    End If
Next
GetStartDataRow = rlt
End Function

Private Function GetEndDataRow(tar As Excel.Range, iRow As Long) As Long
Dim rlt As Long '有数据的第一行，就是起始行
rlt = 0
Dim i As Long
For i = iRow To tar.rows.Count
    If Application.WorksheetFunction.CountA(tar.rows(i)) = 0 Then
        rlt = i - 1
        Exit For
    End If
Next
If rlt = 0 Then '最后一行数据了
    If i >= tar.rows.Count Then rlt = tar.rows.Count
End If

GetEndDataRow = rlt
End Function

Private Function PickFolder() As String
        '** 使用FileDialog对象来选择文件夹
        On Error Resume Next
        Dim fd As FileDialog
        Dim strPath As String
       
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        fd.Title = "选择保存文件夹"
        '** 显示选择文件夹对话框
        If fd.Show = -1 Then        '** 用户选择了文件夹
            strPath = fd.SelectedItems(1)
        Else
            strPath = ""
        End If
        Set fd = Nothing
        PickFolder = strPath
End Function



'WriteTxt "D:\1.txt", "fsdf" & vbTab & "111" & vbCrLf & "1545"
Private Sub WriteTxt(fp As String, txt As String)
        Open fp For Output As #1
        Print #1, txt
        
        'Print #1, "fsdf" & vbTab & "111" & vbCrLf & "1545"
        'print 有自动换行
       ' Write #1, "fsdf" & vbCrLf & "1545"
        Close #1
End Sub