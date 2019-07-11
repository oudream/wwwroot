<%
  Response.Buffer = true
%>

<html>
<head><title>WaitFor组件示例</title></head>
<body>
<h2>WaitFor组件示例</h2>
<blockquote>
<%
  Set WaitObj = Server.CreateObject ("WaitFor.Comp")
  Response.Write "现在开始等待: " & Time & "<br>"
  Response.Flush
  WaitObj.WaitForSeconds 5
  Response.Write "等待完成: " & Time
  Response.Flush
  strFilename = "c:\test.txt"
  Response.Write "<p>现在等待文件 " & strFilename
  Response.Flush
  WaitObj.TimeOut = 45
  if WaitObj.WaitForFileExists (strFileName) then
    Response.Write "<p>在超时时间内找到了文件"
  else
    Response.Write "<p>在超时时间内未找到文件"
  end if
  Response.Flush
%>
</body>
</html>
