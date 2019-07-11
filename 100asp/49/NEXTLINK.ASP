<html>
<head>
<title>
ASP导航器
</title>
</head>
<body>
<center>ASP基础教程章节目录</center>
<ol>
<%  Set NextLink=Server.CreateObject("MSWC.NextLink") %>  
<%  count=NextLink.GetListCount("nextlink.txt") %> 
<%  I=1 %> 
<ul>
<%  Do While (I<=count) %>  
<li><a href=" <%=NextLink.GetNthURL("nextlink.txt", I)  %>  "> 
<%= NextLink.GetNthDescription ("nextlink.txt", I) %>  </a> 
<% I=(I+1)  %>  
<%  Loop  %>  
</ul>
</ol>
</body>
</html> 