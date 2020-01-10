'���=���˳���
'˵��=��˵��
Option Explicit

Sub ��ֱ�������Ϊtxt()
On Error Resume Next
Dim ws As Worksheet, tar As Range, r1 As Range
Set ws = ActiveSheet
Set tar = ws.UsedRange
'�ָ���� ����·��
'�ÿ��д���һ���µ��ļ�
Dim folder As String
folder = PickFolder()
If folder = "" Then Exit Sub
'--------------
Dim sepor As String
Dim tmp As String
tmp = Application.InputBox("��ʲô��Ϊ�ָ��,��ѡ������ѡ�  0:�ո�  1:tab  2:����  ", "�ָ��", "0", Type:=1)
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
Dim rlt As Long '�����ݵĵ�һ�У�������ʼ��
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
Dim rlt As Long '�����ݵĵ�һ�У�������ʼ��
rlt = 0
Dim i As Long
For i = iRow To tar.rows.Count
    If Application.WorksheetFunction.CountA(tar.rows(i)) = 0 Then
        rlt = i - 1
        Exit For
    End If
Next
If rlt = 0 Then '���һ��������
    If i >= tar.rows.Count Then rlt = tar.rows.Count
End If

GetEndDataRow = rlt
End Function

Private Function PickFolder() As String
        '** ʹ��FileDialog������ѡ���ļ���
        On Error Resume Next
        Dim fd As FileDialog
        Dim strPath As String
       
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        fd.Title = "ѡ�񱣴��ļ���"
        '** ��ʾѡ���ļ��жԻ���
        If fd.Show = -1 Then        '** �û�ѡ�����ļ���
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
        'print ���Զ�����
       ' Write #1, "fsdf" & vbCrLf & "1545"
        Close #1
End Sub