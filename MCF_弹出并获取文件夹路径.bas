'类别=工作簿
'说明=

'=================  获取文件夹目录  ==============
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
    Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
    ByVal pszPath As String) As Long

'=================  文件夹选择浏览器  ==============
Public Declare Function SHBrowseForFolder Lib "shell32.dll" _
    Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type


'=====================================================
Sub 弹出并获取文件夹路径()
    
    
    path = GetDirectory("请选择要导入的.bas文件所在的文件夹:")
    If path = "" Then
        Exit Sub
    End If
    
    MsgBox path
    
End Sub

'------弹出并 获取文件夹路径
Function GetDirectory(Optional Msg) As String
    Dim bInfo As BROWSEINFO
    Dim path As String
    Dim r As Long, X As Long, pos As Integer
    ' Root folder = Desktop
    bInfo.pidlRoot = 0&
    
    ' Title in the dialog
    If IsMissing(Msg) Then
        bInfo.lpszTitle = "Select a folder."
    Else
        bInfo.lpszTitle = Msg
    End If
    ' Type of directory to return
    bInfo.ulFlags = &H1
    ' Display the dialog
    X = SHBrowseForFolder(bInfo)
    
    ' Parse the result
    path = Space$(512)
    r = SHGetPathFromIDList(ByVal X, ByVal path)
    If r Then
        pos = InStr(path, Chr$(0))
        GetDirectory = Left(path, pos - 1)
    Else
        GetDirectory = ""
    End If
End Function

