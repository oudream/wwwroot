<%@ LANGUAGE = VBScript %>
<HTML>
<TITLE>
Hello World
</TITLE>
<BODY>
<%
'以下循环输出Hello World字符串,字体由小变大
for i=1 to 5
    response.write "<font size=" & i & ">hello world</font><br>"
next 
%>
</BODY>
</HTML>
