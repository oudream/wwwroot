<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
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
	body=body+"<br>"+"�����ɹ����������Ĳ�����"
end sub

sub sendmsg()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
                <tr> 
                  <th colspan="2">��̳���Ź���
                  </th>
                </tr>
            <form action="admin_message.asp?action=add" method=post>
                <tr> 
                  <td width="22%" class=Forumrow>��Ϣ����</td>
                  <td width="78%" class=Forumrow> 
                    <input type="text" name="title" size="70">
                  </td>
                </tr>
                <tr> 
                  <td width="22%" class=Forumrow>���շ�ѡ��</td>
                  <td width="78%" class=Forumrow> 
                    <select name=stype size=1>
					<option value="1">���������û�</option>
					<option value="2">���й��</option>
					<option value="3">���а���</option>
					<option value="4">���й���Ա</option>
					<option value="5">���/����/����Ա</option>
					<option value="6">�����û�</option>
					</select>
                  </td>
                </tr>
                <tr> 
                  <td width="22%" height="20" valign="top" class=Forumrow>
                    <p>��Ϣ����</p>
                    <p>(<font color="<%=Forum_body(8)%>">HTML����֧��</font>)</p>
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
                      <input type="submit" name="Submit" value="������Ϣ">
                      <input type="reset" name="Submit2" value="������д">
                    </div>
                  </td>
                </tr>
            </form>
                <tr> 
                  <th colspan="2">����ɾ��
                  </th>
                </tr>
            <form action="admin_message.asp?action=del" method=post>
                <tr> 
                  <td colspan="2" class=Forumrow>
                      ����ɾ��ĳ�û�����Ϣ����Ҫ����ɾ��ϵͳ������Ϣ������С���飩��<br><input type="text" name="username" size="20">
			<input type="submit" name="Submit" value="�� ��">
                  </td>
                </tr>
            </form>
			<form action="admin_message.asp?action=delall" method=post>
                <tr> 
                  <td colspan="2" class=Forumrow>
                      ����ɾ���û�ָ�������ڶ���Ϣ��Ĭ��Ϊɾ���Ѷ���Ϣ����<br>
					  <select name="delDate" size=1>
						<option value=7>һ������ǰ</option>
						<option value=30>һ����ǰ</option>
						<option value=60>������ǰ</option>
						<option value=180>����ǰ</option>
						<option value="all">������Ϣ</option>
					  </select>
					  &nbsp;<input type="checkbox" name="isread" value="yes">����δ����Ϣ
			<input type="submit" name="Submit" value="�� ��">
                  </td>
                </tr>
            </form>
			<form action="admin_message.asp?action=delchk" method=post>
                <tr> 
                  <td colspan="2" class=Forumrow>
				  ����ɾ������ĳ�ؼ��ֶ��ţ�ע�⣺��������ɾ�������Ѷ���δ����Ϣ����<br>
				  �ؼ��֣�<input type="text" name="keyword" size=30>&nbsp;��
					  <select name="selaction" size=1>
						<option value=1>������</option>
						<option value=2>������</option>
					  </select>
					  &nbsp;<input type="submit" name="Submit" value="�� ��">
                  </td>
                </tr>
            </form>
              </table>
<%
end sub

sub del()
	if request("username")="" then
		body=body+"<br>"+"������Ҫ����ɾ�����û�����"
		exit sub
	end if
	sql="delete from message where sender='"&request("username")&"'"
	conn.Execute(sql)
	body=body+"<br>"+"�����ɹ����������Ĳ�����"
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
	body=body+"<br>"+"�����ɹ����������Ĳ�����"
end sub

sub delchk()
	if request.form("keyword")="" then
	body="������ؼ��֣�"
	exit sub
	end if
	if request.form("selaction")=1 then
	conn.execute("delete from message where title like '%"&replace(request.form("keyword"),"'","")&"%'")
	body="�����ɹ����������Ĳ�����"
	elseif request.form("selaction")=2 then
	conn.execute("delete from message where content like '%"&replace(request.form("keyword"),"'","")&"%'")
	body="�����ɹ����������Ĳ�����"
	else
	body="δָ����ز�����"
	exit sub
	end if
end sub
%>