Sub PickFileName()
'提取文件名生成文件树或者使用bat命令dir *.*/b>文件清单.txt
On Error GoTo 100

Dim wsh As Object, mypath As String, ar, i&, br

mypath = CreateObject("shell.application").BrowseForFolder(0, "请选择要搜索的文件夹", 0).Items.Item.Path '在此指定目录

Set wsh = CreateObject("wscript.shell")

mypath = wsh.exec("cmd /c tree /f " & Chr(34) & mypath & Chr(34)).StdOut.ReadAll

mypath = Left(mypath, Len(mypath) - 1)

ar = Split(mypath, vbCrLf)

ReDim br(1 To UBound(ar) + 1, 1 To 1)

For i = 0 To UBound(ar)

br(i + 1, 1) = ar(i)

Next

Range("a1").Resize(UBound(br)) = br

Set wsh = Nothing

100:

End Sub