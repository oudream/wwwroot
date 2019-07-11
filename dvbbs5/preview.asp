<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
dim abgcolor
dim username
abgcolor=Forum_body(5)
stats="预览帖子"
%>
<html><head>
<meta NAME=GENERATOR Content="Microsoft FrontPage 4.0" CHARSET=GB2312>
<title><%=Forum_info(0)%>--<%=stats%></title>
<!--#include file="inc/Forum_css.asp"-->
</head>
<body <%=Forum_body(11)%>>
<p>&nbsp;</p>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TBODY> 
<TR>
<Th height=25>帖子预览</Th>
</TR>
<TR> 
<TD class=tablebody1 height=24>
<%
response.write "<b>"& htmlencode(request.form("title")) &"</b><br>"& dvbcode(request.form("body"),4) 
%></TD>
</TR>
</TBODY>
</TABLE>
<%
call footer()
%>