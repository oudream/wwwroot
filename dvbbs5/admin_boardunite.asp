<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	dim str
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		if Request("action") = "unite" then
		call unite()
		else
		call boardinfo()
		end if
		conn.close
		set conn=nothing
	end if

sub boardinfo()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr>
	<th>�ϲ���̳����
	</th>
	</tr>
	<tr>
	<td class=forumrow>ע����� �����棬��������Ŀǰ���е���̳�б�������ͬһ�������ڽ��в����������Խ�һ����̳������һ����̳���кϲ����ϲ���������̳��������ת��ϲ���̳�����ϲ���̳����ɾ���Ҳ��ɻָ���
	</td>
	</tr>
	<tr>
	<td class=forumrow>
<%
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select boardid,boardtype from board"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "û����̳"
	else
		response.write "<form action=admin_boardunite.asp?action=unite method=post>"
		response.write "����̳"
		response.write "<select name=oldboard size=1>"
		do while not rs.eof
		response.write "<option value="&rs("boardid")&">"&rs("boardtype")&"</option>"
		rs.movenext
		loop
		response.write "</select>"
	end if
	rs.close
	sql="select * from board"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "û����̳"
	else
		response.write "�ϲ���"
		response.write "<select name=newboard size=1>"
		do while not rs.eof
		response.write "<option value="&rs("boardid")&">"&rs("boardtype")&"</option>"
		rs.movenext
		loop
		response.write "</select>"
	end if
	rs.close
	set rs=nothing
	response.write "<input type=submit name=Submit value=ִ��></form>"
%>
	</td>
	</tr>
	</table>
<%
end sub

sub unite()
dim newboard
dim oldboard
if cint(request("newboard"))=cint(request("oldboard")) then
response.write "�벻Ҫ����ͬ�����ڽ��в�����"
else
newboard=cint(request("newboard"))
oldboard=cint(request("oldboard"))
'������̳��������
For i=0 to ubound(AllPostTable)
conn.execute("update "&AllPostTable(i)&" set boardid="&newboard&" where boardid="&oldboard)
Next
conn.execute("update topic set boardid="&newboard&" where boardid="&oldboard)
conn.execute("update besttopic set boardid="&newboard&" where boardid="&oldboard)
'��������̳���Ӽ���
set rs=conn.execute("select lastbbsnum,lasttopicnum,todayNum from board where boardid="&oldboard)
conn.execute("update board set lastbbsnum=lastbbsnum+"&rs(0)&",lasttopicnum=lasttopicnum+"&rs(1)&",todaynum=todaynum+"&rs(2)&" where boardid="&newboard)
'ɾ�����ϲ���̳
conn.execute("delete from board where boardid="&oldboard)
response.write "�ϲ��ɹ����Ѿ������ϲ���̳��������ת�������ϲ���̳���뵽������̳����ҳ�������̳���ݡ�"
end if
end sub
%>