'���=
'˵��=��˵��
Option Explicit


Public Sub ��ȡ�ļ��б�()
    Dim ik, iv
    GetFileListArr ThisWorkbook.Path & "\�½��ļ���", "", ik, iv
    
    'MsgBox iv(0)

End Sub

Public Function GetFileListArr(folderspec As String, ext As String, ByRef ik, ByRef iv)

    Dim d As Object
    Set d = CreateObject("scripting.dictionary")
    
    GetFileList folderspec, ext, d
    
    ik = d.keys() '��0��ʼ
    iv = d.items()

End Function



Private Function GetFileList(folderspec As String, ext As String, d As Object) '�������ļ��е�����
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
