<html>
<head><title>ASP实现UNIX加密算法</title></head>
<body>
<h2>ASP实现UNIX加密算法</h2>
<%
  Set Cryptor = Server.CreateObject ("AspCrypt.Crypt")
  strSalt = "a"
  strValue = "Widgets"
  Response.Write "加密前为：" & strValue & "<br>"
  Response.Write "加密后为：" & Cryptor.Crypt (strSalt, strValue)
  Set Cryptor = nothing
%>
</body>
</html>

