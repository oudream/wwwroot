<%
authUser=trim(request.servervariables("AUTH_USER"))
IF authUser="" Then
    Response.Status="401 Not Authorized"
    Response.AddHeader "WWW-Authenticate","Basic Realm="" SUPEREXPERT"""
    Response.End
End if
%>
<Html>
<Head>
<Title>
NT身份认证
</Title>
</Head>
<Body>
<font face=华文彩云 size=5 color=Red><b>
<center>
<%=authUser%>你好!<br>
欢迎你光临我的主页<br>
你已经成功通过身份认证。
</center>
</Body>
</Html>
