'类别=
'说明=无说明

Sub 选择路径()
Dim arr
arr = ChooseMultiFile
End Sub



Public Function ChooseFolder() As String
    On Error Resume Next
    Dim dlgOpen   As FileDialog
    Set dlgOpen = Application.FileDialog(msoFileDialogFolderPicker)
    With dlgOpen
        If .Show = -1 Then
            ChooseFolder = .SelectedItems(1)
        End If
    End With
    
    Set dlgOpen = Nothing
End Function


Public Function ChooseOneFile( _
            Optional TitleStr As String = "选择你要的文件", _
            Optional TypesDec As String = "所有文件", _
            Optional Exten As String = "*.*") As String
            
Dim dlgOpen   As FileDialog
Set dlgOpen = Application.FileDialog(msoFileDialogFilePicker)

With dlgOpen
    .Title = TitleStr
    .Filters.Clear    '清除所有的文件类型.
    .Filters.Add TypesDec, Exten
    .AllowMultiSelect = False    '不能多选.
    
    If .Show = -1 Then
    '                .AllowMultiSelect  =  True              '多个文件
    '                For  Each  vrtSelectedItem  In  .SelectedItems
    '                        MsgBox  Path  name:    &  vrtSelectedItem
    '                Next  vrtSelectedItem
    ChooseOneFile = .SelectedItems(1)          '第一个文件
    End If
End With

End Function

Public Function ChooseMultiFile( _
            Optional TitleStr As String = "选择你要的文件", _
            Optional TypesDec As String = "所有文件", _
            Optional Exten As String = "*.*") As String()
            
Dim dlgOpen   As FileDialog
Set dlgOpen = Application.FileDialog(msoFileDialogFilePicker)

Dim rlt() As String
Dim i As Integer

With dlgOpen
    .Title = TitleStr
    .Filters.Clear    '清除所有的文件类型.
    .Filters.Add TypesDec, Exten
    .AllowMultiSelect = True                '多个文件
    
    If .Show = -1 Then
        ReDim rlt(1 To .SelectedItems.Count)
        For i = 1 To .SelectedItems.Count
            rlt(i) = .SelectedItems(i)
        Next
        
   
    End If
End With

ChooseMultiFile = rlt
End Function
