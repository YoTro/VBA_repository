Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
 
Private Const SW_SHOWNORMAL As Long = 1
 
 
Sub �����ʼ�����()
'�����ʼ�����
ShellExecute 0, "Open", "mailto:zhoujibin123@1126.com", "", "", SW_SHOWNORMAL
'�����������, ���ӵ�Excelhome��̳�������� ShellExecute 0, "Open", _ "http://club.excelhome.net", "", "", SW_SHOWNORMAL
 
End Sub