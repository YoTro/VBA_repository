'���=������
'˵��=��˵��

Option Explicit


Sub �����ļ�()
       On Error Resume Next

    Dim tar As Range
    
    Dim ik, iv
    GetFileListArr ThisWorkbook.Path, "", ik, iv  'ThisWorkbook.PathΪҪ�������ļ���·��
    
    '--------------------------------------------------------
    Set tar = Application.InputBox(prompt:="��ѡ���Ž���ĵ�Ԫ��(����)��", Title:="������", Type:=8)
    
    If tar Is Nothing Then
        Exit Sub
    End If
    
    tar.Cells(1, 1).Offset(0, 0) = "�ļ���"
    tar.Cells(1, 1).Offset(0, 1) = "�ļ�·��"
    
    tar.Cells(1, 1).Offset(1, 0).Resize(UBound(ik) + 1) = WorksheetFunction.Transpose(ik)
    tar.Cells(1, 1).Offset(1, 1).Resize(UBound(iv) + 1) = WorksheetFunction.Transpose(iv)
    Exit Sub

End Sub

Function GetFileListArr(folderspec As String, ext As String, ByRef ik, ByRef iv)

    Dim d As Object
    Set d = CreateObject("scripting.dictionary")
    
    GetFileList folderspec, ext, d
    
    ik = d.keys()
    iv = d.items()

End Function



Public Function GetFileList(folderspec As String, ext As String, d As Object) '�������ļ��е�����
    On Error GoTo l_err
    
    Dim fs
    
    Dim fd As Object
    Dim t As Object
    Dim SerchFD, SerchFF

    
    Set fs = CreateObject("Scripting.FileSystemObject")

    Dim extLen As Integer
    extLen = Len(ext)
    
    Set fd = fs.GetFolder(folderspec)
    Set SerchFD = fd.SubFolders   '�����ļ��ж���
    Set SerchFF = fd.Files       '�����ļ�����
    '�������ļ���
    For Each t In SerchFD
        '�����������ɱ������ļ���
        Call GetFileList(folderspec & "\" & t.Name, ext, d)
    Next
    
    '�����ļ�
    For Each t In SerchFF
        If Right(t.Name, extLen) = ext Then
            d.Add t.Name, folderspec & "\" & t.Name
        End If
    Next

    Exit Function
l_err:
End Function

