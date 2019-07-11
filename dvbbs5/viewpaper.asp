<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
stats="浏览小字报"
%>
<html><head>
<title><%=Forum_info(0)%>--<%=stats%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Forum_css.asp"-->
</head>
<body <%=Forum_body(11)%>>
<%
dim paperid
dim abgcolor
dim username
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的参数。"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的参数。"
else
	paperID=request("id")
end if

if founderr then
	call dvbbs_error()
else
	set rs=server.createobject("adodb.recordset")
	sql="select * from smallpaper where s_id="&paperid
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br>"+"<li>没有找到相关信息。"
	call dvbbs_error()
	else
	conn.execute("update smallpaper set s_hits=s_hits+1 where s_id="&paperid)
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TBODY> 
<TR> 
<Th height=24><%=htmlencode(rs("s_title"))%></Th>
</TR>
<TR> 
<TD class=tablebody1>
<p align=center><a href=dispuser.asp?name=<%=htmlencode(rs("s_username"))%> target=_blank><%=htmlencode(rs("s_username"))%></a> 发布于 <%=rs("s_addtime")%></p>
    <blockquote>   
      <br>   
<%=dvbcode(rs("s_content"),4)%>  
      <br>
<div align=right>浏览次数：<%=rs("s_hits")%></div>
    </blockquote>
</TD>
</TR>
<TR align=middle> 
<TD height=24 class=tablebody2><a href=# onclick="window.close();">『 关闭窗口 』</a></TD>
</TR>
</TBODY>
</TABLE>
<%
	end if
	rs.close
	set rs=nothing
end if
call activeonline()
call footer()
%>
