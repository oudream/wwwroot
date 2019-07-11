<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="md5.asp"-->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
		response.end
	end if
dim trs
dim userinfo
dim usertitle
%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr>
<th align=left colspan=6 height=23>用户管理</th>
</tr>
<tr>
<td width=20% class=forumrow>注意事项</td>
<td width=80% class=forumrow colspan=5>①点删除按钮将删除所选定的用户，此操作是不可逆的；②您可以批量移动用户到相应的组；③点用户名进行相应的资料操作；④点用户最后登陆IP可进行锁定IP操作；⑤点用户Email将给该用户发送Email</td>
</tr>
<form action="?action=userSearch" method=post>
<tr>
<td width=20% class=forumrow>快速搜索</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name="userSearch" onchange="javascript:submit()">
<option value="0">请选择查询条件</option>
<option value="1" <%if request("userSearch")=1 then%>selected<%end if%>>列出所有用户</option>
<option value="2" <%if request("userSearch")=2 then%>selected<%end if%>>发贴最多TOP100</option>
<option value="3" <%if request("userSearch")=3 then%>selected<%end if%>>发贴最少的100个用户</option>
<option value="4" <%if request("userSearch")=4 then%>selected<%end if%>>最近24小时内登陆的用户</option>
<option value="5" <%if request("userSearch")=5 then%>selected<%end if%>>最近24小时内注册的用户</option>
<option value="6" <%if request("userSearch")=6 then%>selected<%end if%>>等待管理员认证的用户</option>
<option value="7" <%if request("userSearch")=7 then%>selected<%end if%>>等待邮件验证的会员</option>
<option value="8" <%if request("userSearch")=8 then%>selected<%end if%>>所有版主组以上用户</option>
</select>
</td>
</tr>
</form>
<%if request("action")="" or request("userSearch")="0" then%>
<form action="?action=userSearch" method=post>
<tr>
<th align=left colspan=6 height=23>高级查询</th>
</tr>
<tr>
<td width=20% class=forumrow>注意事项</td>
<td width=80% class=forumrow colspan=5>在记录很多的情况下搜索条件越多查询越慢，请尽量减少查询条件；最多显示记录数也不宜选择过大</td>
</tr>
<tr>
<td width=20% class=forumrow>最多显示记录数</td>
<td width=80% class=forumrow colspan=5><input size=45 name="searchMax" type=text value=100></td>
</tr>
<tr>
<td width=20% class=forumrow>用户名</td>
<td width=80% class=forumrow colspan=5><input size=45 name="username" type=text>&nbsp;<input type=checkbox name="usernamechk" value="yes" checked>用户名完整匹配</td>
</tr>
<tr>
<td width=20% class=forumrow>用户组</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name="usergroups">
<option value=0>任意</option>
<%
set rs=conn.execute("select usergroupid,title from usergroups order by usergroupid")
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>用户等级</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name="userclass">
<option value=0>任意</option>
<%
set rs=conn.execute("select usertitle from usertitle order by usertitleid")
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(0)&"</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>Email包含</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userEmail" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>主页包含</td>
<td width=80% class=forumrow colspan=5><input size=45 name="homepage" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>QQ包含</td>
<td width=80% class=forumrow colspan=5><input size=45 name="oicq" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>ICQ包含</td>
<td width=80% class=forumrow colspan=5><input size=45 name="icq" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>MSN包含</td>
<td width=80% class=forumrow colspan=5><input size=45 name="msn" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>头衔包含</td>
<td width=80% class=forumrow colspan=5><input size=45 name="usertitle" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>签名包含</td>
<td width=80% class=forumrow colspan=5><input size=45 name="sign" type=text></td>
</tr>
<tr>
<td width=100% class=forumrow align=center colspan=6><input name="submit" type=submit value="   搜  索   "></td>
</tr>
<input type=hidden value="9" name="userSearch">
</form>
<%
elseif request("action")="userSearch" then
%>
<tr>
<th colspan=6 align=left height=23>搜索结果</th>
</tr>
<%
	dim currentpage,page_count,Pcount
	dim totalrec,endpage
	currentPage=request("page")
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
	Set rs= Server.CreateObject("ADODB.Recordset")
	select case request("userSearch")
	case 1
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid order by u.addDate desc"
	case 2
		sql="select top 100  u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid order by u.article desc"
	case 3
		sql="select top 100 u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid order by u.article"
	case 4
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid where datediff('h',u.LastLogin,Now())<25 order by u.lastlogin desc"
	case 5
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid where datediff('h',u.AddDate,Now())<25 order by u.addDate desc"
	case 6
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid where u.usergroupid=5 order by u.addDate desc"
	case 7
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid where u.usergroupid=6 order by u.addDate desc"
	case 8
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid where u.usergroupid<4 order by u.usergroupid"
	case 9
		sqlstr=""
		if request("username")<>"" then
			if request("usernamechk")="yes" then
			sqlstr=" u.username='"&request("username")&"'"
			else
			sqlstr=" u.username like '%"&request("username")&"%'"
			end if
		end if
		if cint(request("usergroups"))>0 then
			if sqlstr="" then
			sqlstr=" u.usergroupid="&request("usergroups")&""
			else
			sqlstr=sqlstr & " and u.usergroupid="&request("usergroupid")&""
			end if
		end if
		if request("userclass")<>"0" then
			if sqlstr="" then
			sqlstr=" u.userclass='"&request("userclass")&"'"
			else
			sqlstr=sqlstr & " and u.userclass='"&request("userclass")&"'"
			end if
		end if
		if request("useremail")<>"" then
			if sqlstr="" then
			sqlstr=" u.useremail like '%"&request("useremail")&"%'"
			else
			sqlstr=sqlstr & " and u.useremail like '%"&request("useremail")&"%'"
			end if
		end if
		if request("homepage")<>"" then
			if sqlstr="" then
			sqlstr=" u.homepage like '%"&request("homepage")&"%'"
			else
			sqlstr=sqlstr & " and u.homepage like '%"&request("homepage")&"%'"
			end if
		end if
		if request("oicq")<>"" then
			if sqlstr="" then
			sqlstr=" u.oicq like '%"&request("oicq")&"%'"
			else
			sqlstr=sqlstr & " and u.oicq like '%"&request("oicq")&"%'"
			end if
		end if
		if request("icq")<>"" then
			if sqlstr="" then
			sqlstr=" u.icq like '%"&request("icq")&"%'"
			else
			sqlstr=sqlstr & " and u.icq like '%"&request("icq")&"%'"
			end if
		end if
		if request("msn")<>"" then
			if sqlstr="" then
			sqlstr=" u.msn like '%"&request("msn")&"%'"
			else
			sqlstr=sqlstr & " and u.msn like '%"&request("msn")&"%'"
			end if
		end if
		if request("title")<>"" then
			if sqlstr="" then
			sqlstr=" u.title like '%"&request("title")&"%'"
			else
			sqlstr=sqlstr & " and u.title like '%"&request("title")&"%'"
			end if
		end if
		if request("sign")<>"" then
			if sqlstr="" then
			sqlstr=" u.sign like '%"&request("sign")&"%'"
			else
			sqlstr=sqlstr & " and u.sign like '%"&request("sign")&"%'"
			end if
		end if
		if sqlstr="" then
		response.write "<tr><td colspan=6 class=forumrow>请指定搜索参数！</td></tr>"
		response.end
		end if
		sql="select top "&request("searchmax")&" u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid where "&sqlstr&" order by u.addDate desc"
	case 10
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.article,u.UserGroupID from [user] u inner join UserGroups G on u.usergroupid=g.usergroupid where u.usergroupid="&request("usergroupid")&" order by u.addDate desc"
	case else
		response.write "<tr><td colspan=6 class=forumrow>错误的参数。</td></tr>"
		response.end
	end select
	'response.write sql
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=6 class=forumrow>没有找到相关记录。</td></tr>"
	else
%>
<FORM METHOD=POST ACTION="?action=touser">
<tr align=center>
<td class=forumHeaderBackgroundAlternate><B>用户名</B></td>
<td class=forumHeaderBackgroundAlternate><B>Email</B></td>
<td class=forumHeaderBackgroundAlternate><B>权限</B></td>
<td class=forumHeaderBackgroundAlternate><B>最后IP</B></td>
<td class=forumHeaderBackgroundAlternate><B>最后登陆</B></td>
<td class=forumHeaderBackgroundAlternate><B>操作</B></td>
</tr>
<%
		rs.PageSize = Cint(Forum_Setting(11))
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = Cint(Forum_Setting(11)))
%>
<tr>
<td class=forumrow><a href="?action=modify&userid=<%=rs("userid")%>"><%=rs("username")%></a></td>
<td class=forumrow width=30% ><a href="mailto:<%=rs("useremail")%>"><%=rs("useremail")%></a></td>
<td class=forumrow width=8% align=center><a href="?action=UserPermission&userid=<%=rs("userid")%>&username=<%=rs("username")%>">编辑</a></td>
<td class=forumrow width=20% ><a href="admin_lockIP.asp?userip=<%=rs("UserLastIP")%>" title="点击锁定该用户IP"><%=rs("userlastip")%></a></td>
<td class=forumrow width=15% ><%if rs("lastlogin")<>"" and isdate(rs("lastlogin")) then%><%=Formatdatetime(rs("lastlogin"),2)%><%end if%></td>
<td class=forumrow align=center><input type="checkbox" name="userid" value="<%=rs("userid")%>" <%if rs("userGroupid")=1 then response.write "disabled"%>></td>
</tr>
<%
		page_count = page_count + 1
		rs.movenext
		wend
Pcount=rs.PageCount
%>
<tr><td colspan=6 class=forumrow align=center>分页：
<%
	if currentpage > 4 then
	response.write "<a href=""?page=1&userSearch="&request("userSearch")&"&username="&request("username")&"&useremail="&request("useremail")&"&homepage="&request("homepage")&"&oicq="&request("oicq")&"&icq="&request("icq")&"&msn="&request("msn")&"&title="&request("title")&"&sign="&request("sign")&"&userclass="&request("userclass")&"&usergroups="&request("usergroups")&"&action="&request("action")&"&usergroupid="&request("usergroupid")&"&searchmax="&request("searchmax")&""">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
	endpage=currentpage+3
	else
	endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color=red>["&i&"]</font>"
		else
        response.write " <a href=""?page="&i&"&userSearch="&request("userSearch")&"&username="&request("username")&"&useremail="&request("useremail")&"&homepage="&request("homepage")&"&oicq="&request("oicq")&"&icq="&request("icq")&"&msn="&request("msn")&"&title="&request("title")&"&sign="&request("sign")&"&userclass="&request("userclass")&"&usergroups="&request("usergroups")&"&action="&request("action")&"&usergroupid="&request("usergroupid")&"&searchmax="&request("searchmax")&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&userSearch="&request("userSearch")&"&username="&request("username")&"&useremail="&request("useremail")&"&homepage="&request("homepage")&"&oicq="&request("oicq")&"&icq="&request("icq")&"&msn="&request("msn")&"&title="&request("title")&"&sign="&request("sign")&"&userclass="&request("userclass")&"&usergroups="&request("usergroups")&"&action="&request("action")&"&usergroupid="&request("usergroupid")&"&searchmax="&request("searchmax")&""">["&Pcount&"]</a>"
	end if
%>
</td></tr>
<tr><td colspan=5 class=forumrow align=center><B>请选择您需要进行的操作</B>：删除<input type="radio" name="useraction" value=1>&nbsp;&nbsp;移动到用户组<input type="radio" name="useraction" value=2 checked>
<select size=1 name="selusergroup">
<%
set trs=conn.execute("select usergroupid,title from usergroups where not (usergroupid=1 or usergroupid=7) order by usergroupid")
do while not trs.eof
response.write "<option value="&trs(0)&">"&trs(1)&"</option>"
trs.movenext
loop
trs.close
set trs=nothing

%>
</select>
</td>
<td class=forumrow align=center><input type=checkbox value="on" name="chkall" onclick="CheckAll(this.form)">
</td>
</tr>
<tr><td colspan=6 class=forumrow align=center>
<input type=submit name=submit value="执行选定的操作"  onclick="{if(confirm('确定执行选择的操作吗?')){this.document.recycle.submit();return true;}return false;}">
</td></tr>
</FORM>
<%
	end if
	rs.close
	set rs=nothing	
elseif request("action")="touser" then
	response.write "<tr><th colspan=6 height=23 align=left>执行结果</th></tr>"
	if request("useraction")="" then
		response.write "<tr><td colspan=6 class=forumrow>请指定相关参数。</td></tr>"
		founderr=true
	end if
	if request("userid")="" then
		response.write "<tr><td colspan=6 class=forumrow>请选择相关用户。</td></tr>"
		founderr=true
	end if
	if not founderr then
		if request("useraction")=1 then
			conn.execute("delete from [user] where userid in ("&replace(request("userid"),"'","")&")")
			response.write "<tr><td colspan=6 class=forumrow>操作成功。</td></tr>"
		elseif request("useraction")=2 then
			conn.execute("update [user] set UserGroupID="&replace(request("selusergroup"),"'","")&" where userid in ("&replace(request("userid"),"'","")&")")
			response.write "<tr><td colspan=6 class=forumrow>操作成功。</td></tr>"
		else
		response.write "<tr><td colspan=6 class=forumrow>错误的参数。</td></tr>"
		end if
	end if
elseif request("action")="modify" then
dim realname,character,personal,country,province,city,shengxiao,blood,belief,occupation,marital, education,college,userphone,iaddress
	response.write "<tr><th colspan=6 height=23 align=left>用户资料操作</th></tr>"
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=6 class=forumrow>错误的用户参数。</td></tr>"
		founderr=true
	end if
	if not founderr then
		Set rs= Server.CreateObject("ADODB.Recordset")
		sql="select * from [user] where userid="&request("userid")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
		response.write "<tr><td colspan=6 class=forumrow>没有找到相关用户。</td></tr>"
		founderr=true
		else
if rs("userinfo")<>"" then
	userinfo=split(rs("userinfo"),"|||")
	if ubound(userinfo)=14 then
		realname=userinfo(0)
		character=userinfo(1)
		personal=userinfo(2)
		country=userinfo(3)
		province=userinfo(4)
		city=userinfo(5)
		shengxiao=userinfo(6)
		blood=userinfo(7)
		belief=userinfo(8)
		occupation=userinfo(9)
		marital=userinfo(10)
		education=userinfo(11)
		college=userinfo(12)
		userphone=userinfo(13)
		iaddress=userinfo(14)
	else
		realname=""
		character=""
		personal=""
		country=""
		province=""
		city=""
		shengxiao=""
		blood=""
		belief=""
		occupation=""
		marital=""
		education=""
		college=""
		userphone=""
		iaddress=""
	end if
else
	realname=""
	character=""
	personal=""
	country=""
	province=""
	city=""
	shengxiao=""
	blood=""
	belief=""
	occupation=""
	marital=""
	education=""
	college=""
	userphone=""
	iaddress=""
end if
%>
<FORM METHOD=POST ACTION="?action=saveuserinfo">
<tr>
<td width=20% class=forumrow valign=top>快捷操作选项</td>
<td width=40% class=forumrow colspan=2 valign=top>
①<a href="mailto:<%=rs("useremail")%>">给 <%=rs("username")%> 发送电子邮件</a><BR>
②<a href="messanger.asp?action=new&touser=<%=rs("username")%>" target=_blank>给 <%=rs("username")%> 发送一条短信</a><BR>
③查找 <%=rs("username")%> 发表的帖子<BR>
④<a href="dispuser.asp?id=<%=rs("userid")%>" target=_blank>预览 <%=rs("username")%> 的显示资料</a>
</td>
<td width=40% class=forumrow colspan=3 valign=top>
⑤<a href="?action=UserPermission&userid=<%=rs("userid")%>&username=<%=rs("username")%>">编辑 <%=rs("username")%> 的论坛权限</a><BR>
⑥查看 <%=rs("username")%> 的最后来源<BR>
⑦<a href="?action=touser&useraction=1&userid=<%=rs("userid")%>">从用户列表删除 <%=rs("username")%></a><BR>
</td>
</tr>
<tr><th colspan=6 height=23 align=left>用户基本资料修改－－<%=rs("username")%></th></tr>
<tr>
<td width=20% class=forumrow>用户组</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name="usergroups">
<%
set trs=conn.execute("select usergroupid,title from usergroups order by usergroupid")
do while not trs.eof
response.write "<option value="&trs(0)
if rs("usergroupid")=trs(0) then response.write " selected "
response.write ">"&trs(1)&"</option>"
trs.movenext
loop
trs.close
set trs=nothing
%>
</select>
</td>
</tr>
<input name="userid" type=hidden value="<%=rs("userid")%>">
<tr>
<td width=20% class=forumrow>用户名</td>
<td width=80% class=forumrow colspan=5><input size=45 name="username" type=text value="<%=rs("username")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>密  码</td>
<td width=80% class=forumrow colspan=5><input size=45 name="password" type=text>&nbsp;如果不修改请留空</td>
</tr>
<tr>
<td width=20% class=forumrow>密码问题</td>
<td width=80% class=forumrow colspan=5><input size=45 name="quesion" type=text value="<%=rs("quesion")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>密码答案</td>
<td width=80% class=forumrow colspan=5><input size=45 name="answer" type=text>&nbsp;如果不修改请留空</td>
</tr>
<tr>
<td width=20% class=forumrow>用户等级</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name="userclass">
<%
set trs=conn.execute("select usertitle from usertitle order by usertitleid")
do while not trs.eof
response.write "<option value="&trs(0)
if rs("userclass")=trs(0) then response.write " selected "
response.write ">"&trs(0)&"</option>"
trs.movenext
loop
trs.close
set trs=nothing
%>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>Email</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userEmail" type=text value="<%=rs("useremail")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>个人主页</td>
<td width=80% class=forumrow colspan=5><input size=45 name="homepage" type=text value="<%=rs("homepage")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>头像</td>
<td width=80% class=forumrow colspan=5><input size=45 name="face" type=text value="<%=rs("face")%>">&nbsp;宽度：<input size=3 name="width" type=text value="<%=rs("width")%>">&nbsp;高度：<input size=3 name="height" type=text value="<%=rs("height")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>OICQ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="oicq" type=text value="<%=rs("oicq")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>ICQ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="icq" type=text value="<%=rs("icq")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>MSN</td>
<td width=80% class=forumrow colspan=5><input size=45 name="msn" type=text value="<%=rs("msn")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>头衔</td>
<td width=80% class=forumrow colspan=5><input size=45 name="usertitle" type=text value="<%=rs("title")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>等级图片</td>
<td width=80% class=forumrow colspan=5><input size=45 name="titlepic" type=text value="<%=rs("titlepic")%>"></td>
</tr>
<tr><th colspan=6 height=23 align=left>用户分值资料修改</th></tr>
<tr>
<td width=20% class=forumrow>发表文章</td>
<td width=80% class=forumrow colspan=5><input size=45 name="article" type=text value="<%=rs("article")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>被删文章</td>
<td width=80% class=forumrow colspan=5><input size=45 name="Userdel" type=text value="<%=rs("userdel")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>精华文章</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userisbest" type=text value="<%=rs("userisbest")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>金钱</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userwealth" type=text value="<%=rs("userwealth")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>经验</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userep" type=text value="<%=rs("userep")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>魅力</td>
<td width=80% class=forumrow colspan=5><input size=45 name="usercp" type=text value="<%=rs("usercp")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>威望</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userpower" type=text value="<%=rs("userpower")%>"></td>
</tr>
<tr><th colspan=6 height=23 align=left>日期相关</th></tr>
<tr>
<td width=20% class=forumrow>生日</td>
<td width=80% class=forumrow colspan=5><input size=45 name="birthday" type=text value="<%=rs("birthday")%>">&nbsp;格式：2001-2-2</td>
</tr>
<tr>
<td width=20% class=forumrow>注册时间</td>
<td width=80% class=forumrow colspan=5><input size=45 name="adddate" type=text value="<%=rs("adddate")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>最后登陆</td>
<td width=80% class=forumrow colspan=5><input size=45 name="lastlogin" type=text value="<%=rs("lastlogin")%>"></td>
</tr>
<tr><th colspan=6 height=23 align=left>用户详细资料</th></tr>
<tr>
<td width=20% class=forumrow>真实姓名</td>
<td width=80% class=forumrow colspan=5><input size=45 name="realname" type=text value="<%=realname%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>国　　家</td>
<td width=80% class=forumrow colspan=5><input size=45 name="country" type=text value="<%=country%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>联系电话</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userphone" type=text value="<%=userphone%>"></td>
</tr><tr>
<td width=20% class=forumrow>通信地址</td>
<td width=80% class=forumrow colspan=5><input size=45 name="address" type=text value="<%=iaddress%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>省　　份</td>
<td width=80% class=forumrow colspan=5><input size=45 name="province" type=text value="<%=province%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>城　　市</td>
<td width=80% class=forumrow colspan=5><input size=45 name="city" type=text value="<%=city%>"></td>
</tr><tr>
<td width=20% class=forumrow>生　　肖</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name=shengxiao>
<option <%if shengxiao="" then%>selected<%end if%>></option>
<option value=鼠 <%if shengxiao="鼠" then%>selected<%end if%>>鼠</option>
<option value=牛 <%if shengxiao="牛" then%>selected<%end if%>>牛</option>
<option value=虎 <%if shengxiao="虎" then%>selected<%end if%>>虎</option>
<option value=兔 <%if shengxiao="兔" then%>selected<%end if%>>兔</option>
<option value=龙 <%if shengxiao="龙" then%>selected<%end if%>>龙</option>
<option value=蛇 <%if shengxiao="蛇" then%>selected<%end if%>>蛇</option>
<option value=马 <%if shengxiao="马" then%>selected<%end if%>>马</option>
<option value=羊 <%if shengxiao="羊" then%>selected<%end if%>>羊</option>
<option value=猴 <%if shengxiao="猴" then%>selected<%end if%>>猴</option>
<option value=鸡 <%if shengxiao="鸡" then%>selected<%end if%>>鸡</option>
<option value=狗 <%if shengxiao="狗" then%>selected<%end if%>>狗</option>
<option value=猪 <%if shengxiao="猪" then%>selected<%end if%>>猪</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>血　　型</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name=blood>
<option <%if blood="" then%>selected<%end if%>></option>
<option value=A <%if blood="A" then%>selected<%end if%>>A</option>
<option value=B <%if blood="B" then%>selected<%end if%>>B</option>
<option value=AB <%if blood="AB" then%>selected<%end if%>>AB</option>
<option value=O <%if blood="O" then%>selected<%end if%>>O</option>
<option value=其他 <%if blood="其他" then%>selected<%end if%>>其他</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>信　　仰</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name=belief>
<option <%if belief="" then%>selected<%end if%>></option>
<option value=佛教 <%if belief="佛教" then%>selected<%end if%>>佛教</option>
<option value=道教 <%if belief="道教" then%>selected<%end if%>>道教</option>
<option value=基督教 <%if belief="基督教" then%>selected<%end if%>>基督教</option>
<option value=天主教 <%if belief="天主教" then%>selected<%end if%>>天主教</option>
<option value=回教 <%if belief="回教" then%>selected<%end if%>>回教</option>
<option value=无神论者 <%if belief="无神论者" then%>selected<%end if%>>无神论者</option>
<option value=共产主义者 <%if belief="共产主义者" then%>selected<%end if%>>共产主义者</option>
<option value=其他 <%if belief="其他" then%>selected<%end if%>>其他</option>
</select>
</td>
</tr><tr>
<td width=20% class=forumrow>职　　业</td>
<td width=80% class=forumrow colspan=5>
<select name=occupation>
<option <%if occupation="" then%>selected<%end if%>> </option>
<option value="财会/金融" <%if occupation="财会/金融" then%>selected<%end if%>>财会/金融</option>
<option value=工程师 <%if occupation="工程师" then%>selected<%end if%>>工程师</option>
<option value=顾问 <%if occupation="顾问" then%>selected<%end if%>>顾问</option>
<option value=计算机相关行业 <%if occupation="计算机相关行业" then%>selected<%end if%>>计算机相关行业</option>
<option value=家庭主妇 <%if occupation="家庭主妇" then%>selected<%end if%>>家庭主妇</option>
<option value="教育/培训" <%if occupation="教育/培训" then%>selected<%end if%>>教育/培训</option>
<option value="客户服务/支持" <%if occupation="客户服务/支持" then%>selected<%end if%>>客户服务/支持</option>
<option value="零售商/手工工人" <%if occupation="零售商/手工工人" then%>selected<%end if%>>零售商/手工工人</option>
<option value=退休 <%if occupation="退休" then%>selected<%end if%>>退休</option>
<option value=无职业 <%if occupation="无职业" then%>selected<%end if%>>无职业</option>
<option value="销售/市场/广告" <%if occupation="销售/市场/广告" then%>selected<%end if%>>销售/市场/广告</option>
<option value=学生 <%if occupation="学生" then%>selected<%end if%>>学生</option>
<option value=研究和开发 <%if occupation="研究和开发" then%>selected<%end if%>>研究和开发</option>
<option value="一般管理/监督" <%if occupation="一般管理/监督" then%>selected<%end if%>>一般管理/监督</option>
<option value="政府/军队" <%if occupation="政府/军队" then%>selected<%end if%>>政府/军队</option>
<option value="执行官/高级管理" <%if occupation="执行官/高级管理" then%>selected<%end if%>>执行官/高级管理</option>
<option value="制造/生产/操作" <%if occupation="制造/生产/操作" then%>selected<%end if%>>制造/生产/操作</option>
<option value=专业人员 <%if occupation="专业人员" then%>selected<%end if%>>专业人员</option>
<option value="自雇/业主" <%if occupation="自雇/业主" then%>selected<%end if%>>自雇/业主</option>
<option value=其他 <%if occupation="其他" then%>selected<%end if%>>其他</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>婚姻状况</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name=marital>
<option <%if marital="" then%>selected<%end if%>></option>
<option value=未婚 <%if marital="未婚" then%>selected<%end if%>>未婚</option>
<option value=已婚 <%if marital="已婚" then%>selected<%end if%>>已婚</option>
<option value=离异 <%if marital="离异" then%>selected<%end if%>>离异</option>
<option value=丧偶 <%if marital="丧偶" then%>selected<%end if%>>丧偶</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>最高学历</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name=education>
<option <%if education="" then%>selected<%end if%>></option>
<option value=小学 <%if education="小学" then%>selected<%end if%>>小学</option>
<option value=初中 <%if education="初中" then%>selected<%end if%>>初中</option>
<option value=高中 <%if education="高中" then%>selected<%end if%>>高中</option>
<option value=大学 <%if education="大学" then%>selected<%end if%>>大学</option>
<option value=硕士 <%if education="硕士" then%>selected<%end if%>>硕士</option>
<option value=博士 <%if education="博士" then%>selected<%end if%>>博士</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>毕业院校</td>
<td width=80% class=forumrow colspan=5><input size=45 name="college" type=text value="<%=college%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>性　格</td>
<td width=80% class=forumrow colspan=5>
<textarea name=character rows=4 cols=80><%=character%></textarea>
</td>
</tr><tr>
<td width=20% class=forumrow>个人简介</td>
<td width=80% class=forumrow colspan=5>
<textarea name=personal rows=4 cols=80><%=personal%></textarea>
</td>
</tr><tr>
<td width=20% class=forumrow>用户签名</td>
<td width=80% class=forumrow colspan=5>
<textarea name="sign" rows=4 cols=80><%=rs("sign")%></textarea>
</td>
</tr>
<tr><th colspan=6 height=23 align=left>用户设置</th></tr>
<tr>
<td width=20% class=forumrow>用户状态</td>
<td width=80% class=forumrow colspan=5>
正常 <input type="radio" value="0" <%if rs("lockuser")=0 then%>checked<%end if%> name="lockuser">&nbsp;
锁定 <input type="radio" value="2" <%if rs("lockuser")=1 then%>checked<%end if%> name="lockuser">&nbsp;
屏蔽 <input type="radio" value="3" <%if rs("lockuser")=2 then%>checked<%end if%> name="lockuser">
</td>
</tr>
<tr>
<td width=100% class=forumrow align=center colspan=6><input name="submit" type=submit value="   更  新   "></td>
</tr>
</FORM>
<%
		end if
		rs.close
		set rs=nothing
	end if
elseif request("action")="saveuserinfo" then
	response.write "<tr><th colspan=6 height=23 align=left>更新用户资料</th></tr>"
	userinfo=checkreal(request.Form("realname")) & "|||" & checkreal(request.Form("character")) & "|||" & checkreal(request.Form("personal")) & "|||" & checkreal(request.Form("country")) & "|||" & checkreal(request.Form("province")) & "|||" & checkreal(request.Form("city")) & "|||" & request.Form("shengxiao") & "|||" & request.Form("blood") & "|||" & request.Form("belief") & "|||" & request.Form("occupation") & "|||" & request.Form("marital") & "|||" & request.Form("education") & "|||" & checkreal(request.Form("college")) & "|||" & checkreal(request.Form("userphone")) & "|||" & checkreal(request.Form("address"))
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=6 class=forumrow>错误的用户参数。</td></tr>"
		founderr=true
	end if
	if not founderr then
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select * from [user] where userid="&request("userid")
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=6 class=forumrow>没有找到相关用户。</td></tr>"
		founderr=true
	else
		rs("username")=request.form("username")
		if request.form("password")<>"" then
		rs("userpassword")=md5(request.form("password"))
		end if
		rs("usergroupid")=request.form("usergroups")
		rs("quesion")=request.form("quesion")
		if request.form("answer")<>"" then
		rs("answer")=md5(request.form("answer"))
		end if
		rs("userclass")=request.form("userclass")
		rs("useremail")=request.form("useremail")
		rs("homepage")=request.form("homepage")
		rs("face")=request.form("face")
		rs("width")=request.form("width")
		rs("height")=request.form("height")
		rs("oicq")=request.form("oicq")
		rs("icq")=request.form("icq")
		rs("msn")=request.form("msn")
		rs("title")=request.form("usertitle")
		rs("titlepic")=request.form("titlepic")
		rs("article")=request.form("article")
		rs("userdel")=request.form("userdel")
		rs("userisbest")=request.form("userisbest")
		rs("userpower")=request.form("userpower")
		rs("userwealth")=request.form("userwealth")
		rs("userep")=request.form("userep")
		rs("usercp")=request.form("usercp")
		rs("birthday")=request.form("birthday")
		rs("addDate")=request.form("adddate")
		rs("lastlogin")=request.form("lastlogin")
		rs("lockuser")=request.form("lockuser")
		rs("sign")=request.form("sign")
		rs("userinfo")=userinfo
		rs.update
	end if
	rs.close
	set rs=nothing
	end if
	if founderr then
		response.write "<tr><td colspan=6 class=forumrow>更新失败。</td></tr>"
	else
		response.write "<tr><td colspan=6 class=forumrow>更新用户数据成功。</td></tr>"
	end if
elseif request("action")="UserPermission" then
	response.write "<tr><th colspan=6 height=23 align=left>编辑"&request("username")&"论坛权限（红色表示该用户在该版面有自定义权限）</th></tr>"
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=6 class=forumrow>错误的用户参数。</td></tr>"
		founderr=true
	end if
	if not founderr then
	response.write "<tr><td colspan=6 class=forumrow height=25><a href=?action=userBoardPermission&boardid=0&userid="&request("userid")&">编辑该用户在其它页面的权限</a>（主要针对短信部分设置）</td></tr>"
'----------------------boardinfo--------------------
response.write "<tr><td colspan=6 class=forumrow>"
dim rs_1,rs_2
dim sql_1,sql_2
	sql_2 = "select * from class order by orders"
	set rs_2=conn.execute(sql_2)
	do while not rs_2.Eof
%>
            <table width="95%" border="0" cellspacing="3" cellpadding="0" align=center>
              <tr>

                <td height="21"><B><%=rs_2("class")%></b></td>
              </tr>
            </table>
<%
		sql_1 = "select boardid,boardtype from board where class="&rs_2("id")&" order by orders"
		set rs_1=conn.execute(sql_1)
		do while not rs_1.EOF 
%>
            <table width="95%" border="0" cellspacing="3" cellpadding="0" align=center>
              <tr> 
                <td height="18" width=50% ><li><%=rs_1("boardtype")%></td>
				<td width=50% >
<a href="?action=userBoardPermission&boardid=<%=rs_1("boardid")%>&userid=<%=request("userid")%>">
<%
set trs=conn.execute("select uc_userid from UserAccess where uc_Boardid="&rs_1("boardid")&" and uc_userid="&request("userid"))
if not (trs.eof and trs.bof) then
	response.write "<font color=red>"
end if
%>
编辑用户在该论坛权限</a></td>
              </tr>
            </table>
<%
		rs_1.MoveNext
		loop
%>
            <table width="95%" border="0" cellspacing="3" cellpadding="0" align=center>
              <tr>
                <td height="21" align=left><hr color=black height=1 width=70% size=1></td>
              </tr>
            </table>
<%
	rs_2.MoveNext 
	Loop
	set trs=nothing
set rs_1=nothing
set rs_2=nothing
response.write "</td></tr>"
'-------------------end-------------------
	end if
elseif request("action")="userBoardPermission" then
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=6 class=forumrow>错误的用户参数。</td></tr>"
		founderr=true
	end if
	if not isnumeric(request("boardid")) then
		response.write "<tr><td colspan=6 class=forumrow>错误的版面参数。</td></tr>"
		founderr=true
	end if
	if not founderr then
	set rs=conn.execute("select u.UserGroupID,ug.title,u.username from [user] u inner join UserGroups UG on u.userGroupID=ug.userGroupID where u.userid="&request("userid"))
	UserGroupID=rs(0)
	usertitle=rs(1)
	membername=rs(2)
	set rs=conn.execute("select boardtype from board where boardid="&request("boardid"))
	if rs.eof and rs.bof then
	boardtype="论坛其他页面"
	else
	boardtype=rs(0)
	end if
	response.write "<tr><th colspan=6 height=23 align=left>编辑 "&membername&" 在 "&boardtype&" 权限</th></tr>"
	response.write "<tr><td colspan=6 height=25 class=forumrow>注意：该用户属于 <B>"&usertitle&"</B> 用户组中，如果您设置了他的自定义权限，则该用户权限将以自定义权限为主</td></tr>"
%>
<tr><td colspan=6 class=forumrow>
<%
Dim reGroupSetting
Dim FoundGroup,FoundUserPermission,FoundGroupPermission
FoundGroup=false
FoundUserPermission=false
FoundGroupPermission=false

set rs=conn.execute("select * from UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
if not (rs.eof and rs.bof) then
	reGroupSetting=split(rs("uc_Setting"),",")
	FoundGroup=true
	FoundUserPermission=true
end if

if not foundgroup then
set rs=conn.execute("select * from BoardPermission where boardid="&request("boardid")&" and groupid="&usergroupid)
if not(rs.eof and rs.bof) then
	reGroupSetting=split(rs("PSetting"),",")
	FoundGroup=true
	FoundGroupPermission=true
end if
end if

if not foundgroup then
set rs=conn.execute("select * from usergroups where usergroupid="&usergroupid)
if rs.eof and rs.bof then
	response.write "未找到该用户组！"
	response.end
else
	FoundGroup=true
	FoundGroupPermission=true
	reGroupSetting=split(rs("GroupSetting"),",")
end if
end if
%>
<table width="100%" border="0" cellspacing="1" cellpadding="0"  align=center>
<FORM METHOD=POST ACTION="?action=saveuserpermission">
<input type=hidden name="userid" value="<%=request("userid")%>">
<input type=hidden name="BoardID" value="<%=request("boardid")%>">
<input type=hidden name="username" value="<%=membername%>">
<tr> 
<td height="23" colspan="2" class=forumrow><input type=radio name="isdefault" value="1" <%if FoundGroupPermission then%>checked<%end if%>><B>使用用户组默认值</B> (注意: 这将删除任何之前所做的自定义设置)</td>
</tr>
<tr> 
<td height="23" colspan="2"  class=forumrow><input type=radio name="isdefault" value="0" <%if FoundUserPermission then%>checked<%end if%>><B>使用自定义设置</B> </td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝查看权限</th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以浏览论坛</td>
<td height="23" width="40%" class=forumrow>是<input name="canview" type=radio value="1" <%if reGroupSetting(0)=1 then%>checked<%end if%>>&nbsp;否<input name="canview" type=radio value="0" <%if reGroupSetting(0)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以查看会员信息(包括其他会员的资料和会员列表)
</td>
<td height="23" width="40%" class=forumrow>是<input name="canviewuserinfo" type=radio value="1" <%if reGroupSetting(1)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewuserinfo" type=radio value="0" <%if reGroupSetting(1)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以查看其他人发布的主题
</td>
<td height="23" width="40%" class=forumrow>是<input name="canviewpost" type=radio value="1" <%if reGroupSetting(2)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewpost" type=radio value="0" <%if reGroupSetting(2)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以浏览精华帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canviewbest" type=radio value="1" <%if reGroupSetting(41)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewbest" type=radio value="0" <%if reGroupSetting(41)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝<b>发帖权限</b></th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以发布新主题</td>
<td height="23" width="40%" class=forumrow>是<input name="cannewpost" type=radio value="1" <%if reGroupSetting(3)=1 then%>checked<%end if%>>&nbsp;否<input name="cannewpost" type=radio value="0" <%if reGroupSetting(3)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以回复自己的主题
</td>
<td height="23" width="40%" class=forumrow>是<input name="canreplymytopic" type=radio value="1" <%if reGroupSetting(4)=1 then%>checked<%end if%>>&nbsp;否<input name="canreplymytopic" type=radio value="0" <%if reGroupSetting(4)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以回复其他人的主题
</td>
<td height="23" width="40%" class=forumrow>是<input name="canreplytopic" type=radio value="1" <%if reGroupSetting(5)=1 then%>checked<%end if%>>&nbsp;否<input name="canreplytopic" type=radio value="0" <%if reGroupSetting(5)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以在论坛允许评分的时候参与评分(鲜花和鸡蛋)?
</td>
<td height="23" width="40%" class=forumrow>是<input name="canpostagree" type=radio value="1" <%if reGroupSetting(6)=1 then%>checked<%end if%>>&nbsp;否<input name="canpostagree" type=radio value="0" <%if reGroupSetting(6)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>参与评分所需金钱
</td>
<td height="23" width="40%" class=Forumrow><input name="postagreemoney" type=text size=4 value="<%=reGroupSetting(47)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以上传附件
</td>
<td height="23" width="40%" class=forumrow>是<input name="canupload" type=radio value="1" <%if reGroupSetting(7)=1 then%>checked<%end if%>>&nbsp;否<input name="canupload" type=radio value="0" <%if reGroupSetting(7)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>最多上传文件个数
</td>
<td height="23" width="40%" class=Forumrow><input name="canuploadnum" type=text size=4 value="<%=reGroupSetting(40)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>上传文件大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="MaxUploadSize" type=text size=4 value="<%=reGroupSetting(44)%>"> KB</td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以发布新投票</td>
<td height="23" width="40%" class=forumrow>是<input name="canpostvote" type=radio value="1" <%if reGroupSetting(8)=1 then%>checked<%end if%>>&nbsp;否<input name="canpostvote" type=radio value="0" <%if reGroupSetting(8)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以参与投票</td>
<td height="23" width="40%" class=forumrow>是<input name="canvote" type=radio value="1" <%if reGroupSetting(9)=1 then%>checked<%end if%>>&nbsp;否<input name="canvote" type=radio value="0" <%if reGroupSetting(9)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布小字报</td>
<td height="23" width="40%" class=Forumrow>是<input name="cansmallpaper" type=radio value="1"  <%if reGroupSetting(17)=1 then%>checked<%end if%>>&nbsp;否<input name="cansmallpaper" type=radio value="0" <%if reGroupSetting(17)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>发布小字报所需金钱</td>
<td height="23" width="40%" class=Forumrow><input name="smallpapermoney" type=text value="<%=reGroupSetting(46)%>" size=4></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝<b>帖子/主题编辑权限</b></th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以编辑自己的帖子
</td>
<td height="23" width="40%" class=forumrow>是<input name="caneditmytopic" type=radio value="1" <%if reGroupSetting(10)=1 then%>checked<%end if%>>&nbsp;否<input name="caneditmytopic" type=radio value="0" <%if reGroupSetting(10)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以删除自己的帖子
</td>
<td height="23" width="40%" class=forumrow>是<input name="candelmytopic" type=radio value="1" <%if reGroupSetting(11)=1 then%>checked<%end if%>>&nbsp;否<input name="candelmytopic" type=radio value="0" <%if reGroupSetting(11)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以移动自己的帖子到其他论坛
</td>
<td height="23" width="40%" class=forumrow>是<input name="canmovemytopic" type=radio value="1" <%if reGroupSetting(12)=1 then%>checked<%end if%>>&nbsp;否<input name="canmovemytopic" type=radio value="0" <%if reGroupSetting(12)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以打开/关闭自己发布的主题
</td>
<td height="23" width="40%" class=forumrow>是<input name="canclosemytopic" type=radio value="1" <%if reGroupSetting(13)=1 then%>checked<%end if%>>&nbsp;否<input name="canclosemytopic" type=radio value="0" <%if reGroupSetting(13)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝<b>其他权限</b></th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以搜索论坛
</td>
<td height="23" width="40%" class=forumrow>是<input name="cansearch" type=radio value="1" <%if reGroupSetting(14)=1 then%>checked<%end if%>>&nbsp;否<input name="cansearch" type=radio value="0" <%if reGroupSetting(14)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以使用'发送本页给好友'功能
</td>
<td height="23" width="40%" class=forumrow>是<input name="canmailtopic" type=radio value="1" <%if reGroupSetting(15)=1 then%>checked<%end if%>>&nbsp;否<input name="canmailtopic" type=radio value="0" <%if reGroupSetting(15)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>可以修改个人资料
</td>
<td height="23" width="40%" class=forumrow>是<input name="canmodify" type=radio value="1" <%if reGroupSetting(16)=1 then%>checked<%end if%>>&nbsp;否<input name="canmodify" type=radio value="0" <%if reGroupSetting(16)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>帖子中可以查看其它人头衔
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canusetitle" type=radio value="1" <%if reGroupSetting(36)=1 then%>checked<%end if%>>&nbsp;否<input name="canusetitle" type=radio value="0" <%if reGroupSetting(36)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>帖子中可以查看其它人头像
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canuseface" type=radio value="1"  <%if reGroupSetting(37)=1 then%>checked<%end if%>>&nbsp;否<input name="canuseface" type=radio value="0" <%if reGroupSetting(37)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>帖子中可以可以查看其它人签名
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canusesign" type=radio value="1"  <%if reGroupSetting(38)=1 then%>checked<%end if%>>&nbsp;否<input name="canusesign" type=radio value="0" <%if reGroupSetting(38)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以浏览论坛事件
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canvieweven" type=radio value="1"  <%if reGroupSetting(39)=1 then%>checked<%end if%>>&nbsp;否<input name="canvieweven" type=radio value="0" <%if reGroupSetting(39)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝管理权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以删除其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="candeltopic" type=radio value="1" <%if reGroupSetting(18)=1 then%>checked<%end if%>>&nbsp;否<input name="candeltopic" type=radio value="0"  <%if reGroupSetting(18)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以移动其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmovetopic" type=radio value="1" <%if reGroupSetting(19)=1 then%>checked<%end if%>>&nbsp;否<input name="canmovetopic" type=radio value="0"  <%if reGroupSetting(19)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以打开/关闭其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canclosetopic" type=radio value="1" <%if reGroupSetting(20)=1 then%>checked<%end if%>>&nbsp;否<input name="canclosetopic" type=radio value="0"  <%if reGroupSetting(20)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以固顶/解除固顶帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="cantoptopic" type=radio value="1" <%if reGroupSetting(21)=1 then%>checked<%end if%>>&nbsp;否<input name="cantoptopic" type=radio value="0"  <%if reGroupSetting(21)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以奖励/惩罚发贴用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canawardtopic" type=radio value="1" <%if reGroupSetting(22)=1 then%>checked<%end if%>>&nbsp;否<input name="canawardtopic" type=radio value="0"  <%if reGroupSetting(22)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以奖励/惩罚用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canaward" type=radio value="1" <%if reGroupSetting(43)=1 then%>checked<%end if%>>&nbsp;否<input name="canaward" type=radio value="0"  <%if reGroupSetting(43)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以编辑其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmodifytopic" type=radio value="1" <%if reGroupSetting(23)=1 then%>checked<%end if%>>&nbsp;否<input name="canmodifytopic" type=radio value="0" <%if reGroupSetting(23)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以加入/解除精华帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canbesttopic" type=radio value="1" <%if reGroupSetting(24)=1 then%>checked<%end if%>>&nbsp;否<input name="canbesttopic" type=radio value="0"  <%if reGroupSetting(24)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布公告
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAnnounce" type=radio value="1" <%if reGroupSetting(25)=1 then%>checked<%end if%>>&nbsp;否<input name="canAnnounce" type=radio value="0"  <%if reGroupSetting(25)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理公告
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAdminAnnounce" type=radio value="1" <%if reGroupSetting(26)=1 then%>checked<%end if%>>&nbsp;否<input name="canAdminAnnounce" type=radio value="0"  <%if reGroupSetting(26)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理小字报
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAdminPaper" type=radio value="1" <%if reGroupSetting(27)=1 then%>checked<%end if%>>&nbsp;否<input name="canAdminPaper" type=radio value="0"  <%if reGroupSetting(27)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以锁定/屏蔽/解除锁定用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAdminUser" type=radio value="1" <%if reGroupSetting(28)=1 then%>checked<%end if%>>&nbsp;否<input name="canAdminUser" type=radio value="0"  <%if reGroupSetting(28)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以删除用户1－10天内所发帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canDelUserTopic" type=radio value="1" <%if reGroupSetting(29)=1 then%>checked<%end if%>>&nbsp;否<input name="canDelUserTopic" type=radio value="0"  <%if reGroupSetting(29)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以查看来访IP及来源
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canviewip" type=radio value="1" <%if reGroupSetting(30)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewip" type=radio value="0"  <%if reGroupSetting(30)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以限定IP来访
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canadminip" type=radio value="1" <%if reGroupSetting(31)=1 then%>checked<%end if%>>&nbsp;否<input name="canadminip" type=radio value="0"  <%if reGroupSetting(31)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理用户权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="adminpermission" type=radio value="1" <%if reGroupSetting(42)=1 then%>checked<%end if%>>&nbsp;否<input name="adminpermission" type=radio value="0"  <%if reGroupSetting(42)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以批量删除帖子（前台）
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canbatchtopic" type=radio value="1" <%if reGroupSetting(45)=1 then%>checked<%end if%>>&nbsp;否<input name="canbatchtopic" type=radio value="0"  <%if reGroupSetting(45)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝短信权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发送短信
</td>
<td height="23" width="40%" class=Forumrow>是<input name="cansendsms" type=radio value="1"  <%if reGroupSetting(32)=1 then%>checked<%end if%>>&nbsp;否<input name="cansendsms" type=radio value="0" <%if reGroupSetting(32)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>最多发送用户
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsendsms" size=5 type=text value="<%=reGroupSetting(33)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>短信内容大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbody" size=5 type=text value="<%=reGroupSetting(34)%>"> byte</td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>信箱大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbox" size=5 type=text value="<%=reGroupSetting(35)%>"> KB</td>
</tr>
<tr> 
<td height="23" width="60%" class=forumrow>
</td>
<td height="23" width="40%" class=forumrow><input type="submit" name="submit" value="提 交"></td>
</tr>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
</td></tr>
<%
	end if
elseif request("action")="saveuserpermission" then
	response.write "<tr><th colspan=6 height=23 align=left>编辑用户 "&request("username")&" 权限</th></tr>"
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=6 class=forumrow>错误的用户参数。</td></tr>"
		founderr=true
	end if
	if not isnumeric(request("boardid")) then
		response.write "<tr><td colspan=6 class=forumrow>错误的版面参数。</td></tr>"
		founderr=true
	end if
	if not founderr then
	dim myGroupSetting
	myGroupSetting=Request.Form("canview") & "," & Request.Form("canviewuserinfo") & "," & Request.Form("canviewpost") & "," & Request.Form("cannewpost") & "," & Request.Form("canreplymytopic") & "," & Request.Form("canreplytopic") & "," & Request.Form("canpostagree") & "," & Request.Form("canupload") & "," & Request.Form("canpostvote") & "," & Request.Form("canvote") & "," & Request.Form("caneditmytopic") & "," & Request.Form("candelmytopic") & "," & Request.Form("canmovemytopic") & "," & Request.Form("canclosemytopic") & "," & Request.Form("cansearch") & "," & Request.Form("canmailtopic") & "," & Request.Form("canmodify") & "," & Request.Form("cansmallpaper") & "," & Request.Form("candeltopic") & "," & Request.Form("canmovetopic") & "," & Request.Form("canclosetopic") & "," & Request.Form("cantoptopic") & "," & Request.Form("canawardtopic") & "," & Request.Form("canmodifytopic") & "," & Request.Form("canbesttopic") & "," & Request.Form("canAnnounce") & "," & Request.Form("canAdminAnnounce") & "," & Request.Form("canAdminPaper") & "," & Request.Form("canAdminUser") & "," & Request.Form("canDelUserTopic") & "," & Request.Form("canviewip") & "," & Request.Form("canadminip") & "," & Request.Form("cansendsms") & "," & Request.Form("Maxsendsms") & "," & Request.Form("Maxsmsbody") & "," & Request.Form("Maxsmsbox") & "," & Request.Form("canusetitle") & "," & Request.Form("canuseface") & "," & Request.Form("canusesign") & "," & Request.Form("canvieweven") & "," & Request.Form("canuploadnum") & "," & Request.Form("canviewbest") & "," & Request.Form("adminpermission") & "," & request.form("canaward") & "," & request.form("MaxUploadSize") & "," & request.form("canbatchtopic") & "," & request.form("smallpapermoney") & "," & request.form("postagreemoney")
	if request("isdefault")=1 then
		conn.execute("delete from UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
	else
		set rs=conn.execute("select * from UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
		if rs.eof and rs.bof then
			conn.execute("insert into UserAccess (uc_userid,uc_boardid,uc_setting) values ("&request("userid")&","&request("boardid")&",'"&myGroupSetting&"')")
		else
			conn.execute("update UserAccess set uc_setting='"&myGroupSetting&"' where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
		end if
	end if
	if founderr then
		response.write "<tr><td colspan=6 class=forumrow>更新失败。</td></tr>"
	else
		response.write "<tr><td colspan=6 class=forumrow>设置用户权限成功。</td></tr>"
	end if
	end if
end if
function checkreal(v)
	dim w
	if not isnull(v) then
	w=replace(v,"|||","§§§")
	checkreal=w
	end if
end function
conn.close
set conn=nothing
%>
</table>
<p></p>

<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>