<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<%
dim voteid
dim title
dim voteoption
stats="�鿴ͶƱ�û�"
if request("id")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
	founderr=true
elseif not isInteger(request("id")) then
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
	founderr=true
else
	VoteID=request("id")
end if
if founderr then
	call dvbbs_error()
	response.end
else
	set rs=conn.execute("select title from topic where pollid="&voteid)
	if not (rs.eof and rs.bof) then
	title=rs(0)
	end if
end if
%>
<html><head>
<title><%=Forum_info(0)%>--<%=stats%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Forum_css.asp"-->
</head>
<body <%=Forum_body(11)%>>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TBODY> 
<TR> 
<Th height=24 colspan=2><%=title%></Th>
</TR>
<TR align=center> 
<TD class=tablebody2><B>�û�</B>
</TD>
<TD class=tablebody2><B>��Ŀ(���ֱ�ʾ�ڼ���)</B>
</TD>
</TR>
<%
set rs=conn.execute("select v.*,u.username from voteuser v inner join [user] u on v.userid=u.userid where voteid="&voteid)
if rs.eof and rs.bof then
%>
<TR> 
<TD class=tablebody1 colspan=2>��û����ͶƱ��
</TD>
</TR>
<%
else
do while not rs.eof
%>
<TR> 
<TD class=tablebody1><a href="dispuser.asp?id=<%=rs("userid")%>" target=_blank><%=rs("username")%></a>
</TD>
<TD class=tablebody1>
<%
voteoption=split(rs("voteoption"),",")
for i=0 to ubound(voteoption)
if isnumeric(voteoption(i)) then response.write voteoption(i)+1
if i<>ubound(voteoption) then response.write ","
next
%>
<%'=rs("voteoption")%>
</TD>
</TR>
<%
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
</TBODY>
</TABLE>
<%
call activeonline()
conn.close
set conn=nothing
%>