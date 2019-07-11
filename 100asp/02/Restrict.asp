<%@ Language=VBScript %>
<html>
<title>
本站主页
</title>
<body>
<%
'本例根据远程主机地址来进行判断,如果为本地
'地址则进入欢迎页面,否则显示出错信息.
dim address
address = request.servervariables("REMOTE_ADDR")
if address="127.0.0.1" then 
    '如果为本地,则显示欢迎页面.
    response.write "你好,欢迎进入本站点."
else
    '否则显示出错信息.
    response.write "对不起,你无权查看内部站点."
end if
%>
</body>
</html>