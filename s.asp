<html>
<head>
<title>send to your email</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="Jmail" content="use jmail send to your email">
<meta name="send to your email" content="use jmail send to your email">
<title>send to your email</title>
</head>

<body>
<%
'下面是定义一些变量，是这些变量从表单页面获取信息
company=Request.Form ("company")
website=Request.Form ("website")
fname=Request.Form ("fname")
lname=Request.Form ("lname")
email=Request.Form ("email")
phone=Request.Form ("phone")
country=Request.Form ("country")
city=Request.Form ("city")
message=Request.Form ("message")

' 下面就是调用从表单页获取的信息，赋值到mess，&是连接符，vbcrlf表示换行回车:

mess = mess & "---------------Webmail表单开始-------------------" & vbcrlf
mess = mess & "Company:" & company & vbcrlf
mess = mess & "Website:" & website & vbcrlf
mess = mess & "First name:" & fname & vbcrlf
mess = mess & "Last name:" & lname & vbcrlf
mess = mess & "Email address:" & email & vbcrlf
mess = mess & "Phone:" & phone & vbcrlf
mess = mess & "Country:" & country & vbcrlf 
mess = mess & "City / Locality:" & city & vbcrlf
mess = mess & "Message:" & message & vbcrlf

mess = mess & "---------------Webmail表单结束-------------------" & vbcrlf
' 下面是调用Jmail组件的部分，其中邮箱，密码填写正确的，否则发送不成功
Set JMail=Server.CreateObject("JMail.Message")
JMail.silent = true                 '则errorcode包含的是错误代码
JMail.Logging=True                '是否使用日志
JMail.Charset="gb2312"
JMail.ContentType = "text/plain"       'text/html是超文本格式text/plain是文本格式
JMail.MailServerUserName = "2264496741@qq.com" '您的邮件服务器登录名
JMail.MailServerPassword = "aaayuye990210"         '登录密码
JMail.From = "2264496741@qq.com"                  '发件人(要填合法正确邮箱才能发出去信件)
JMail.FromName = "gewell"                 '发件人姓名，引号内的内容可以修改为符合自己的信息
JMail.AddRecipient "2264496741@qq.com"      '收件人
JMail.Subject = "WebMail"      '主题，引号内的内容可以修改为自己认为合适的主题
JMail.Body = mess     '正文，通过使用前面的mess调用获取到得表单信息

JMail.Send "mail.qq.com"                        'smtp服务器地址
JMail.Priority=3                               '1: 最高优先级.2:高 3:普通 4:低 5:最低
Set JMail=nothing
%>
<CENTER>
Success ，Congratulation
</CENTER>
<p align="center">--- &lt; <a href="index.htm">Go back to your website </a>&gt; ---</p> 
</body>

</html>