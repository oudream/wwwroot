<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	Server.ScriptTimeout=9999999
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		dim body
		call main()
		conn.close
		set conn=nothing
	end if

	sub main()
	select case request("action")
	case "alldel"
		call alldel()
	case "userdel"
		call del()
	case "alldelTopic"
		call alldelTopic()
	case "deluser"
		call deluser()
	case "moveinfo"
		call moveinfo()
	case "MoveUserTopic"
		call moveusertopic()
	case "MoveDateTopic"
		call movedatetopic()
	case else
%>
<table cellpadding=3 cellspacing=1 border=0 width=95% align=center>
	<tr>
    <td width="100%" valign=top>
<B>注意</B>：如果您想还原帖子，请到论坛回收站！
            <br>下面操作将大批量删除论坛帖子。如果您确定这样做，请仔细检查您输入的信息。
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<form action="admin_alldel.asp?action=alldel" method="post">
            <tr>
            <th valign=middle colspan=2 height=23 align=left>删除指定日期内帖子</b>(本功能不扣除用户帖子数和积分)</th></tr>
            <tr>
            <td valign=middle width=40% class=forumrow>删除多少天前的帖子(填写数字)</td><td class=forumrow><input name="TimeLimited" value=0 size=30>&nbsp;<input type=submit name="submit" value="提 交"></td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>论坛版面</td><td class=forumrow>
<select name="delboardid" size=1>
<%
set rs=conn.execute("select boardid,boardtype from board order by boardid")
if rs.eof and rs.bof then
response.write "<option value=0>没有论坛</option>"
else
response.write "<option value=all>全部论坛</option>"
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
</select>
			</td></tr>
</form>
<form action="admin_alldel.asp?action=alldelTopic" method="post">
            <tr>
            <th valign=middle colspan=2 height=23 align=left>删除指定日期内没有回复的主题(本功能不扣除用户帖子数和积分)</th></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>删除多少天前的帖子(填写数字)</td><td class=forumrow><input name="TimeLimited" value=0 size=30>&nbsp;<input type=submit name="submit" value="提 交"></td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>论坛版面</td><td class=forumrow>
<select name="delboardid" size=1>
<%
set rs=conn.execute("select boardid,boardtype from board order by boardid")
if rs.eof and rs.bof then
response.write "<option value=0>没有论坛</option>"
else
response.write "<option value=all>全部论坛</option>"
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
</select>
			</td></tr>
</form>
<form action="admin_alldel.asp?action=userdel" method="post">
            <tr>
            <th valign=middle colspan=2 height=23 align=left>删除某用户的所有帖子</td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>请输入用户名</td><td class=forumrow><input type=text name="username" size=30>&nbsp;<input type=submit name="submit" value="提 交"></td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>论坛版面</td><td class=forumrow>
<select name="delboardid" size=1>
<%
set rs=conn.execute("select boardid,boardtype from board order by boardid")
if rs.eof and rs.bof then
response.write "<option value=0>没有论坛</option>"
else
response.write "<option value=all>全部论坛</option>"
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
</select>
			</td></tr>
</form>
<!--
<form action="admin_alldel.asp?action=delUser" method="post">
            <tr>
            <td bgcolor="<%=Forum_body(4)%>" valign=middle><font color="<%=Forum_body(7)%>">删除指定日期内没有登陆的用户</font></td>
            <td bgcolor="<%=Forum_body(4)%>" valign=middle>
<select name=TimeLimited size=1> 
<option value=all>删除所有的
<option value=1>删除一天前的
<option value=2>删除两天前的
<option value=7>删除一星期前的
<option value=15>删除半个月前的
<option value=30>删除一个月前的
<option value=60>删除两个月前的
<option value=180>删除半年前的
</select>
</select><input type=submit name="submit" value="提 交"></td></tr></form>
-->
</table>
<%end select%>
<%if founderr then call dvbbs_error()%>
<%
	end sub

	sub moveinfo()
%>
<table cellpadding=3 cellspacing=1 border=0 width=95% align=center>
	<tr>
    <td width="100%" valign=top>
<B>注意</B>：这里只是移动帖子，而不是拷贝或者删除！
            <br>下面操作将删除原论坛帖子，并移动到您指定的论坛中。如果您确定这样做，请仔细检查您输入的信息。
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<form action="admin_alldel.asp?action=MoveDateTopic" method="post">
            <tr>
            <th valign=middle colspan=2 height=23 align=left>按日期移动</th></tr>
            <tr>
            <td valign=middle width=40% class=forumrow>移动多少天前的帖子(填写数字)</td><td class=forumrow><input name="TimeLimited" value=0 size=30>&nbsp;<input type=submit name="submit" value="提 交"></td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>原论坛</td><td class=forumrow>
<select name="outboardid" size=1>
<%
set rs=conn.execute("select boardid,boardtype from board order by boardid")
if rs.eof and rs.bof then
response.write "<option value=0>没有论坛</option>"
else
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
</select>
			</td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>目标论坛</td><td class=forumrow>
<select name="inboardid" size=1>
<%
set rs=conn.execute("select boardid,boardtype from board order by boardid")
if rs.eof and rs.bof then
response.write "<option value=0>没有论坛</option>"
else
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
</select>
			</td></tr>
</form>
<form action="admin_alldel.asp?action=MoveUserTopic" method="post">
            <tr>
            <th valign=middle colspan=2 height=23 align=left>按用户移动</th></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>请填写用户名</td><td class=forumrow><input name="username" size=30>&nbsp;<input type=submit name="submit" value="提 交"></td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>原论坛</td><td class=forumrow>
<select name="outboardid" size=1>
<%
set rs=conn.execute("select boardid,boardtype from board order by boardid")
if rs.eof and rs.bof then
response.write "<option value=0>没有论坛</option>"
else
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
</select>
			</td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>目标论坛</td><td class=forumrow>
<select name="inboardid" size=1>
<%
set rs=conn.execute("select boardid,boardtype from board order by boardid")
if rs.eof and rs.bof then
response.write "<option value=0>没有论坛</option>"
else
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
</select>
			</td></tr>
</form>
</table>
<%
	end sub
	sub del()
		dim titlenum,delboardid
		if request("delboardid")="0" then
			founderr=true
			errmsg=errmsg+"<br>"+"<li>非法的版面参数。"
			exit sub
		elseif request("delboardid")="all" then
			delboardid=""
		else
			delboardid=" boardid="&request("delboardid")&" and "
		end if
		if request("username")="" then
			founderr=true
			errmsg=errmsg+"<br>"+"<li>请输入被帖子删除用户名。"
			exit sub
		end if
		titlenum=0
		for i=0 to ubound(allposttable)
		set rs=conn.execute("Select Count(announceID) from "&allposttable(i)&" where "&delboardid&"   username='"&replace(request("username"),"'","")&"'") 
   		titlenum=titlenum+rs(0)

		sql="update "&allposttable(i)&" set locktopic=2 where "&delboardid&"  username='"&replace(request("username"),"'","")&"'"
		conn.Execute(sql)
		next
		set rs=conn.execute("select topicid,posttable from topic where "&delboardid&"  postusername='"&replace(request("username"),"'","")&"'")
		do while not rs.eof
		conn.execute("update "&rs(1)&" set locktopic=2 where rootid="&rs(0))
		rs.movenext
		loop
		set rs=nothing
		conn.execute("update topic set locktopic=2 where "&delboardid&"  Postusername='"&replace(request("username"),"'","")&"'")
		if isnull(titlenum) then titlenum=0
		sql="update [user] set article=article-"&titlenum&",userWealth=userWealth-"&titlenum*Forum_user(3)&",userEP=userEP-"&titlenum*Forum_user(8)&",userCP=userCP-"&titlenum*Forum_user(13)&" where username='"&replace(request("username"),"'","")&"'"
		conn.Execute(sql)
	response.write "删除成功，如果要完全删除帖子请到论坛回收站<BR>建议您到更新论坛数据中更新一下论坛数据，或者<a href=admin_alldel.asp>返回</a>"
	end sub

	sub alldel()
	Dim TimeLimited,delboardid
	if request("delboardid")="0" then
		founderr=true
		errmsg=errmsg+"<br>"+"<li>非法的版面参数。"
		exit sub
	elseif request("delboardid")="all" then
		delboardid=""
	else
		delboardid=" and boardid="&request("delboardid")&" "
	end if
	TimeLimited=request.form("TimeLimited")
	if not isnumeric(TimeLimited) then
		founderr=true
		errmsg=errmsg+"<br>"+"<li>非法的参数。"
		exit sub
	else
	for i=0 to ubound(allposttable)
	conn.execute("update "&allposttable(i)&" set LockTopic=2 where datediff('d',DateAndTime,Now())>"&TimeLimited&" "&delboardid&"")
	next
	conn.execute("update topic set LockTopic=2 where datediff('d',DateAndTime,Now())>"&TimeLimited&" "&delboardid&"")
	end if
	response.write "删除成功，如果要完全删除帖子请到论坛回收站<BR>建议您到更新论坛数据中更新一下论坛数据，或者<a href=admin_alldel.asp>返回</a>"
	end sub

	sub alldelTopic()
	Dim TimeLimited,delboardid
	if request("delboardid")="0" then
		founderr=true
		errmsg=errmsg+"<br>"+"<li>非法的版面参数。"
		exit sub
	elseif request("delboardid")="all" then
		delboardid=""
	else
		delboardid=" boardid="&request("delboardid")&" and "
	end if
	TimeLimited=request.form("TimeLimited")
	if not isnumeric(TimeLimited) then
		founderr=true
		errmsg=errmsg+"<br>"+"<li>非法的参数。"
		exit sub
	else
	conn.execute("update topic set LockTopic=2 where "&delboardid&"   datediff('d',DateAndTime,Now())>"&TimeLimited&" and Child=0")
	set rs=conn.execute("select Topicid,PostTable from topic where "&delboardid&"   datediff('d',DateAndTime,Now())>"&TimeLimited&" and Child=0")
	do while not rs.eof
		conn.execute("update "&rs(1)&" set locktopic=2 where rootid="&rs(0))
	rs.movenext
	loop
	set rs=nothing
	end if
	response.write "删除成功，如果要完全删除帖子请到论坛回收站<BR>建议您到更新论坛数据中更新一下论坛数据，或者<a href=admin_alldel.asp>返回</a>"
	end sub

	sub delUser()
	Dim TimeLimited
	TimeLimited=request.form("TimeLimited")
	if TimeLimited="all" then
	conn.execute("delete from [user]")
	else
	conn.execute("delete from [user] where datediff('d',LastLogin,Now())>"&TimeLimited&"")
	end if
	response.write "删除成功，如果要完全删除帖子请到论坛回收站<BR>建议您到更新论坛数据中更新一下论坛数据，或者<a href=admin_alldel.asp>返回</a>"
	end sub

	sub MoveUserTopic()
	if not isnumeric(request("inboardid")) then
	response.write "错误的版面参数。"
	exit sub
	end if
	if not isnumeric(request("outboardid")) then
	response.write "错误的版面参数。"
	exit sub
	end if
	if request("username")="" then
	response.write "请填写用户名。"
	exit sub
	end if
	if Cint(request("outboardid"))=Cint(request("inboardid")) then
	response.write "不能在相同版面进行移动操作！"
	exit sub
	end if
	for i=0 to ubound(allposttable)
	conn.execute("update "&allposttable(i)&" set boardid="&request("inboardid")&" where Boardid="&request("outboardid")&" and username='"&request("username")&"'")
	next
	set rs=conn.execute("select topicid,posttable from topic where Boardid="&request("outboardid")&" and Postusername='"&request("username")&"'")
	do while not rs.eof
		conn.execute("update "&rs(1)&" set boardid="&request("inboardid")&" where rootid="&rs(0))
	rs.movenext
	loop
	set rs=nothing
	conn.execute("update topic set boardid="&request("inboardid")&" where Boardid="&request("outboardid")&" and Postusername='"&request("username")&"'")
	conn.execute("update besttopic set boardid="&request("inboardid")&" where Boardid="&request("outboardid")&" and Postusername='"&request("username")&"'")
	response.write "移动成功！"
	end sub

	sub MoveDateTopic()
	if not isnumeric(request("TimeLimited")) then
	response.write "错误的日期参数。"
	exit sub
	end if
	if not isnumeric(request("inboardid")) then
	response.write "错误的版面参数。"
	exit sub
	end if
	if not isnumeric(request("outboardid")) then
	response.write "错误的版面参数。"
	exit sub
	end if
	if Cint(request("outboardid"))=Cint(request("inboardid")) then
	response.write "不能在相同版面进行移动操作！"
	exit sub
	end if
	for i=0 to ubound(allposttable)
	conn.execute("update "&allposttable(i)&" set boardid="&request("inboardid")&" where Boardid="&request("outboardid")&" and datediff('d',DateAndTime,Now())>"&request.Form("TimeLimited")&"")
	next
	conn.execute("update topic set boardid="&request("inboardid")&" where Boardid="&request("outboardid")&" and datediff('d',DateAndTime,Now())>"&request.Form("TimeLimited")&"")
	conn.execute("update besttopic set boardid="&request("inboardid")&" where Boardid="&request("outboardid")&" and datediff('d',DateAndTime,Now())>"&request.Form("TimeLimited")&"")
	response.write "移动成功！"
	end sub
%>