<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		dim body
		if Request("action")="add" then
			call savemsg()
		elseif request("action")="del" then
			call del()
		elseif request("action")="delall" then
			call delall()
		elseif request("action")="delchk" then
			call delchk()
		else
			call sendmsg()
		end if
%>
<p align=center><%=body%></p></font>
<%
		conn.close
		set conn=nothing
	end if

sub savemsg()
	dim sendtime,sender
	sendtime=Now()
	sender=Forum_info(0)
	set rs = server.CreateObject ("adodb.recordset")
	select case request("stype")
	case 1
	sql="select username from online where userid>0"
	rs.open sql,conn,1,1
	do while not rs.eof
	sql="insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&rs(0)&"','"&sender&"','"&TRim(Request("title"))&"','"&Trim(Request("message"))&"',Now(),0,1)"
	conn.Execute(sql)
	rs.movenext
	loop
	rs.close
	case 2
    sql = "select username from [user] where usergroupid=8 order by userid desc"
    rs.Open sql,conn,1,1
    do while not rs.EOF 

	sql="insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&rs(0)&"','"&sender&"','"&TRim(Request("title"))&"','"&Trim(Request("message"))&"',Now(),0,1)"
	conn.Execute(sql)
	rs.MoveNext 
	Loop
	rs.Close
	case 3
    sql = "select username from [user] where usergroupid=3 order by userid desc"
    rs.Open sql,conn,1,1
    do while not rs.EOF 

	sql="insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&rs(0)&"','"&sender&"','"&TRim(Request("title"))&"','"&Trim(Request("message"))&"',Now(),0,1)"
	conn.Execute(sql)
	rs.MoveNext 
	Loop
	rs.Close
	case 4
    sql = "select username from [user] where usergroupid=1 order by userid desc"
    rs.Open sql,conn,1,1
    do while not rs.EOF 

	sql="insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&rs(0)&"','"&sender&"','"&TRim(Request("title"))&"','"&Trim(Request("message"))&"',Now(),0,1)"
	conn.Execute(sql)
	rs.MoveNext 
	Loop
	rs.Close
	case 5
    sql = "select username from [user] where usergroupid<4 order by userid desc"
    rs.Open sql,conn,1,1
    do while not rs.EOF 

	sql="insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&rs(0)&"','"&sender&"','"&TRim(Request("title"))&"','"&Trim(Request("message"))&"',Now(),0,1)"
	conn.Execute(sql)
	rs.MoveNext 
	Loop
	rs.Close
	case 6
    sql = "select username from [user] order by userid desc"
    rs.Open sql,conn,1,1
    do while not rs.EOF 

	sql="insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&rs(0)&"','"&sender&"','"&TRim(Request("title"))&"','"&Trim(Request("message"))&"',Now(),0,1)"
	conn.Execute(sql)
	rs.MoveNext 
	Loop
	rs.Close
	end select
	set rs=nothing
	body=body+"<br>"+"操作成功！请继续别的操作。"
end sub

sub sendmsg()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
                <tr> 
                  <th colspan="2">论坛短信管理
                  </th>
                </tr>
            <form action="admin_message.asp?action=add" method=post>
                <tr> 
                  <td width="22%" class=Forumrow>消息标题</td>
                  <td width="78%" class=Forumrow> 
                    <input type="text" name="title" size="70">
                  </td>
                </tr>
                <tr> 
                  <td width="22%" class=Forumrow>接收方选择</td>
                  <td width="78%" class=Forumrow> 
                    <select name=stype size=1>
					<option value="1">所有在线用户</option>
					<option value="2">所有贵宾</option>
					<option value="3">所有版主</option>
					<option value="4">所有管理员</option>
					<option value="5">贵宾/版主/管理员</option>
					<option value="6">所有用户</option>
					</select>
                  </td>
                </tr>
                <tr> 
                  <td width="22%" height="20" valign="top" class=Forumrow>
                    <p>消息内容</p>
                    <p>(<font color="<%=Forum_body(8)%>">HTML代码支持</font>)</p>
                  </td>
                  <td width="78%" height="20" class=Forumrow> 
                    <textarea name="message" cols="80" rows="10"></textarea>
                  </td>
                </tr>
                <tr> 
                  <td width="22%" height="23" valign="top" align="center" class=Forumrow> 
                    <div align="left"> </div>
                  </td>
                  <td width="78%" height="23" class=Forumrow> 
                    <div align="center"> 
                      <input type="submit" name="Submit" value="发送消息">
                      <input type="reset" name="Submit2" value="重新填写">
                    </div>
                  </td>
                </tr>
            </form>
                <tr> 
                  <th colspan="2">批量删除
                  </th>
                </tr>
            <form action="admin_message.asp?action=del" method=post>
                <tr> 
                  <td colspan="2" class=Forumrow>
                      批量删除某用户短消息（主要用于删除系统批量信息：动网小精灵）：<br><input type="text" name="username" size="20">
			<input type="submit" name="Submit" value="提 交">
                  </td>
                </tr>
            </form>
			<form action="admin_message.asp?action=delall" method=post>
                <tr> 
                  <td colspan="2" class=Forumrow>
                      批量删除用户指定日期内短消息（默认为删除已读信息）：<br>
					  <select name="delDate" size=1>
						<option value=7>一个星期前</option>
						<option value=30>一个月前</option>
						<option value=60>两个月前</option>
						<option value=180>半年前</option>
						<option value="all">所有信息</option>
					  </select>
					  &nbsp;<input type="checkbox" name="isread" value="yes">包括未读信息
			<input type="submit" name="Submit" value="提 交">
                  </td>
                </tr>
            </form>
			<form action="admin_message.asp?action=delchk" method=post>
                <tr> 
                  <td colspan="2" class=Forumrow>
				  批量删除含有某关键字短信（注意：本操作将删除所有已读和未读信息）：<br>
				  关键字：<input type="text" name="keyword" size=30>&nbsp;在
					  <select name="selaction" size=1>
						<option value=1>标题中</option>
						<option value=2>内容中</option>
					  </select>
					  &nbsp;<input type="submit" name="Submit" value="提 交">
                  </td>
                </tr>
            </form>
              </table>
<%
end sub

sub del()
	if request("username")="" then
		body=body+"<br>"+"请输入要批量删除的用户名。"
		exit sub
	end if
	sql="delete from message where sender='"&request("username")&"'"
	conn.Execute(sql)
	body=body+"<br>"+"操作成功！请继续别的操作。"
end sub

sub delall()
	dim selflag
	if request("isread")="yes" then
	selflag=""
	else
	selflag=" and flag=1"
	end if
	select case request("delDate")
	case "all"
	sql="delete from message where id>0 "&selflag
	case 7
	sql="delete from message where datediff('d',sendtime,Now())>7 "&selflag
	case 30
	sql="delete from message where datediff('d',sendtime,Now())>30 "&selflag
	case 60
	sql="delete from message where datediff('d',sendtime,Now())>60 "&selflag
	case 180
	sql="delete from message where datediff('d',sendtime,Now())>180 "&selflag
	end select
	conn.Execute(sql)
	body=body+"<br>"+"操作成功！请继续别的操作。"
end sub

sub delchk()
	if request.form("keyword")="" then
	body="请输入关键字！"
	exit sub
	end if
	if request.form("selaction")=1 then
	conn.execute("delete from message where title like '%"&replace(request.form("keyword"),"'","")&"%'")
	body="操作成功！请继续别的操作。"
	elseif request.form("selaction")=2 then
	conn.execute("delete from message where content like '%"&replace(request.form("keyword"),"'","")&"%'")
	body="操作成功！请继续别的操作。"
	else
	body="未指定相关参数！"
	exit sub
	end if
end sub
%>