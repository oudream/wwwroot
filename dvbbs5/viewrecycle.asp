<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
stats="����վ����"
	dim AnnounceID
	dim username
	dim tablename
	dim abgcolor
	if request("id")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ��Ĳ�����"
	elseif not isInteger(request("id")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ��Ĳ�����"
	else
		AnnounceID=request("id")
	end if
	if not master then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>������ϵͳ����Ա��û��Ȩ�޲鿴��ɾ�������ӡ�"
	end if
	if founderr then
		call dvbbs_error(1)
	else
		set rs=server.createobject("adodb.recordset")
		if request("tablename")="bbs1" then
			sql="select Announceid,topic,body,username,dateandtime from bbs1 where AnnounceID="&AnnounceID
			tablename=request("tablename")
		elseif request("tablename")="bbs2" then
			sql="select Announceid,topic,body,username,dateandtime from bbs2 where AnnounceID="&AnnounceID
			tablename=request("tablename")
		else
			sql="select topicid,title,title as body,postusername,dateandtime from topic where topicID="&AnnounceID
			tablename="topic"
		end if
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br>"+"<li>û���ҵ������Ϣ��"
		call dvbbs_error(1)
		else
%>
<html><head>
<meta NAME=GENERATOR Content="Microsoft FrontPage 4.0" CHARSET=GB2312>
<title><%=Forum_info(0)%>--<%=stats%></title>
<!--#include file=inc/forum_css.asp-->
</head>
<body <%=Forum_body(11)%>>
<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1>
<TBODY> 
<TR align=middle> 
<Th height=24><%=htmlencode(rs(1))%></Th>
</TR>
<TR> 
<TD height=24 class=tablebody1>
<p align=center><a href=dispuser.asp?name=<%=htmlencode(rs(3))%> target=_blank><%=htmlencode(rs(3))%></a> ������ <%=rs(4)%></p>
    <blockquote>   
      <br>   
<%=dvbcode(rs(2),1)%>
    </blockquote>
</TD>
</TR>
<TR align=middle> 
<TD height=24 class=tablebody2><a href="admin_recycle.asp?TopicID=<%=rs(0)%>&action=ɾ��&tablename=<%=tablename%>">�� ֱ��ɾ�� ��</a>&nbsp;&nbsp;<a href=# onclick="window.close();">�� �رմ��� ��</a></TD>
</TR>
</TBODY>
</TABLE>
    </td>
  </tr>
</table>
<%
		end if
		rs.close
		set rs=nothing
	end if
	call footer()
%>
