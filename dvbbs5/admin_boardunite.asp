<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	dim str
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
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
	<th>合并论坛数据
	</th>
	</tr>
	<tr>
	<td class=forumrow>注意事项： 在下面，您将看到目前所有的论坛列表。不能在同一个版面内进行操作。您可以将一个论坛和另外一个论坛进行合并，合并后两个论坛所有帖子转入合并论坛，被合并论坛将被删除且不可恢复。
	</td>
	</tr>
	<tr>
	<td class=forumrow>
<%
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select boardid,boardtype from board"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "没有论坛"
	else
		response.write "<form action=admin_boardunite.asp?action=unite method=post>"
		response.write "将论坛"
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
		response.write "没有论坛"
	else
		response.write "合并到"
		response.write "<select name=newboard size=1>"
		do while not rs.eof
		response.write "<option value="&rs("boardid")&">"&rs("boardtype")&"</option>"
		rs.movenext
		loop
		response.write "</select>"
	end if
	rs.close
	set rs=nothing
	response.write "<input type=submit name=Submit value=执行></form>"
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
response.write "请不要在相同版面内进行操作。"
else
newboard=cint(request("newboard"))
oldboard=cint(request("oldboard"))
'更新论坛帖子数据
For i=0 to ubound(AllPostTable)
conn.execute("update "&AllPostTable(i)&" set boardid="&newboard&" where boardid="&oldboard)
Next
conn.execute("update topic set boardid="&newboard&" where boardid="&oldboard)
conn.execute("update besttopic set boardid="&newboard&" where boardid="&oldboard)
'更新新论坛帖子计数
set rs=conn.execute("select lastbbsnum,lasttopicnum,todayNum from board where boardid="&oldboard)
conn.execute("update board set lastbbsnum=lastbbsnum+"&rs(0)&",lasttopicnum=lasttopicnum+"&rs(1)&",todaynum=todaynum+"&rs(2)&" where boardid="&newboard)
'删除被合并论坛
conn.execute("delete from board where boardid="&oldboard)
response.write "合并成功，已经将被合并论坛所有数据转入您所合并论坛，请到更新论坛数据页面更新论坛数据。"
end if
end sub
%>