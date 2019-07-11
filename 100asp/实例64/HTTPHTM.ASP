<html>
<head>
<title>HTTPhtm</title>
</head>
<body>
<%
  Server.ScriptTimeout = 240
  Set HttpObj = Server.CreateObject("AspHTTP.Conn")
  HttpObj.Url = "http://localhost/test/test.htm"
  rem HttpObj.Port = 80
  rem HttpObj.TimeOut = 80
  HttpObj.FollowRedirects = true
  HttpObj.RequestMethod = "GET"
  HttpObj.UserAgent = "Mozilla/2.0 (compatible; MSIE 3.0B; Windows NT)"
  strResult = HttpObj.GetURL
  Response.Write strResult
  Response.Write "<blockquote><hr>URL is now: " & HttpObj.URL
  Response.Write "<hr><pre>" & HttpObj.Headers & "</pre>"
  Response.Write "<hr>"
  strRegName = HttpObj.RegisteredUser
%>
</body>
</html>