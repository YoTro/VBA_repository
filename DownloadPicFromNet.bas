Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szExtName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
'批量下载网上的图片，电脑必须是64位
Sub DownloadPictures()

     Dim strKey As String

     Dim strURL As String

     Dim strFolderPath As String

     Dim strText As String

     Dim strPicPath As String

     Dim strPicURL As String

     Dim strExtName As String

     Dim aPageNum As Variant

     Dim aExtName As Variant

     Dim i As Long

     Dim k As Long

     strFolderPath = ThisWorkbook.Path & "\图片\"

     If Dir(strFolderPath, vbDirectory + vbHidden) > "" Then

         If Dir(strFolderPath & "*.*") > "" Then Kill strFolderPath & "*.*"

     Else

         MkDir strFolderPath

     End If

     strKey = [a2].Value

     If Len(strKey) = 0 Then

         MsgBox "未输入查询关键字，程序退出。"

         Exit Sub

     End If

     strKey = encodeURI(strKey) '对查询关键字转码

     With CreateObject("msxml2.xmlhttp") '发送网页请求，获得响应信息

         strURL = "http://image.baidu.com/search/index?tn=baiduimage&word=" & strKey

         .Open "GET", strURL, "False"

         .send

         strText = .responseText

     End With

     aPageNum = Split(strText, """pageNum"":")

     '按关键字pageNum对响应信息进行拆分

     For i = 1 To UBound(aPageNum)

         If InStr(1, aPageNum(i), "objURL", vbTextCompare) Then

         '判断是否存在字符串objurl

             k = k + 1

             strPicURL = Split(Split(aPageNum(i), """objURL"":""")(1), """,")(0)

             '图片的网址

             aExtName = Split(strPicURL, ".")

             strExtName = "." & aExtName(UBound(aExtName))

             '图片的后缀名

             strPicPath = strFolderPath & k & strExtName

             '图片保存地址

             DeleteUrlCacheEntry strPicURL

             '删除图片缓存数据

             URLDownloadToFile 0, strPicURL, strPicPath, 0, 0

             '下载图片

         End If

     Next

End Sub

Function encodeURI(strText As String) As String

    Dim objDOM As Object

    Set objDOM = CreateObject("htmlfile")

    With objDOM.parentWindow

        objDOM.Write "<Script></Script>"

        encodeURI = .eval("encodeURIComponent('" & strText & "')")

    End With

    Set objDOM = Nothing

End Function