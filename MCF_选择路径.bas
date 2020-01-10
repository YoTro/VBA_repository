'���=
'˵��=��˵��

Sub ѡ��·��()
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
            Optional TitleStr As String = "ѡ����Ҫ���ļ�", _
            Optional TypesDec As String = "�����ļ�", _
            Optional Exten As String = "*.*") As String
            
Dim dlgOpen   As FileDialog
Set dlgOpen = Application.FileDialog(msoFileDialogFilePicker)

With dlgOpen
    .Title = TitleStr
    .Filters.Clear    '������е��ļ�����.
    .Filters.Add TypesDec, Exten
    .AllowMultiSelect = False    '���ܶ�ѡ.
    
    If .Show = -1 Then
    '                .AllowMultiSelect  =  True              '����ļ�
    '                For  Each  vrtSelectedItem  In  .SelectedItems
    '                        MsgBox  Path  name:    &  vrtSelectedItem
    '                Next  vrtSelectedItem
    ChooseOneFile = .SelectedItems(1)          '��һ���ļ�
    End If
End With

End Function

Public Function ChooseMultiFile( _
            Optional TitleStr As String = "ѡ����Ҫ���ļ�", _
            Optional TypesDec As String = "�����ļ�", _
            Optional Exten As String = "*.*") As String()
            
Dim dlgOpen   As FileDialog
Set dlgOpen = Application.FileDialog(msoFileDialogFilePicker)

Dim rlt() As String
Dim i As Integer

With dlgOpen
    .Title = TitleStr
    .Filters.Clear    '������е��ļ�����.
    .Filters.Add TypesDec, Exten
    .AllowMultiSelect = True                '����ļ�
    
    If .Show = -1 Then
        ReDim rlt(1 To .SelectedItems.Count)
        For i = 1 To .SelectedItems.Count
            rlt(i) = .SelectedItems(i)
        Next
        
   
    End If
End With

ChooseMultiFile = rlt
End Function
