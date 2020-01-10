Sub CDOsendMail()
'通过excel从qq邮箱中发送邮件
'1，获取qq邮箱的smtp服务码方式如下。

'打开网页版qq邮箱，依次单击【设置】→【账户】；找到POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV服务，选择开启IMAP/SMTP服务。开启SMTP服务后，会获得相关密码。
'2，.Item(strURL & "sendusername") = strFromName

'上述代码设置的发件人账户名称，是“账户设置”表B3单元格的值，该值只是账户名称，比如469772827，不是邮箱地址，比如469772827@qq.com

'3，变量strPath指定了邮件添加附件存放的路径和名称，如果需要给不同的人发送不同的附件请参阅文末列出的往期推文。

'4，如果将一封邮件发送多人，不同收件人之间使用半角分号间隔即可，例如：

'"46@qq.com;47@qq.com;48@qq.com"

'5，代码稍加修改也可以用于使用163邮箱发送邮件。修改发件人的邮箱地址、名称和对应的smtp服务密码。同时将以下语句：

'.Item(strURL & "smtpserver") = "smtp.qq.com"

'修改为：.Item(strURL & "smtpserver") = "smtp.163.com"

   Dim CDOMail As Object
    Dim strPath As String
    Dim aData As Variant
    Dim i As Long
    Dim strURL As String
    Dim strFromMail As String
    Dim strFromName As String
    Dim strPassWord As String
    strFromMail = Range("b2").Value
    strFromName = Range("b3").Value
    If strFromMail = "" Or strFromName = "" Then
        MsgBox "未输入邮箱地址或名称。"
        Exit Sub
    End If
    strPassWord = Range("b4").Value
    If strPassWord = "****" Or strPassWord = "" Then
        MsgBox "未输入smtp服务密码"
        Exit Sub
    End If
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    Sheets("数据").Select
    aData = Range("a1:d" & Cells(Rows.Count, 1).End(xlUp).Row)
    '--------数据装入数组aData
    strPath = ThisWorkbook.Path & "/暑假快乐.jpg"
    '--------附件路径
    On Error Resume Next
    For i = 2 To UBound(aData)
        Set CDOMail = CreateObject("CDO.Message")
    '--------创建CDO对象
        CDOMail.From = strFromMail
    '--------发信人的邮箱
        CDOMail.To = aData(i, 1)
    '--------收信人的邮箱
        CDOMail.Subject = aData(i, 2)
    '--------邮件的主题
        CDOMail.HtmlBody = aData(i, 3)
    '--------邮件的内容（Html格式)
        'CDOMail.TextBody = aData(i, 3)
    '--------邮件的内容（文本格式)
        CDOMail.AddAttachment strPath
    '--------邮件的附件
        strURL = "http://schemas.microsoft.com/cdo/configuration/"
    '--------微软服务器网址
        With CDOMail.Configuration.Fields
            .Item(strURL & "smtpserver") = "smtp.qq.com"
    '--------SMTP服务器地址
            .Item(strURL & "smtpserverport") = 25
    '--------SMTP服务器端口
            .Item(strURL & "sendusing") = 2
    '--------发送端口
            .Item(strURL & "smtpauthenticate") = 1
    '--------远程服务器验证
            .Item(strURL & "sendusername") = strFromName
    '--------发送方邮箱名称
            .Item(strURL & "sendpassword") = strPassWord
    '--------发送方smtp密码
            .Item(strURL & "smtpconnectiontimeout") = 60
    '--------设置连接超时（秒）
            .Update
        End With
        CDOMail.Send
    '--------发送
        If Err.Number = 0 Then
            aData(i, 1) = "发送成功"
        Else
            aData(i, 1) = "发送失败"
        End If
    Next
    Range("d1").Resize(UBound(aData), 1) = aData
    Range("d1") = "发送状态"
    Set CDOMail = Nothing
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    MsgBox "您好，发送任务完成。"
End Sub

