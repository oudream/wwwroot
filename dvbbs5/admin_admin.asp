<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="md5.asp"-->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<script language="JavaScript">
<!--
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall')
       e.checked = form.chkall.checked;
    }
  }
//-->
</script>
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		dim body,username2,password2,oldpassword,oldusername,oldadduser,username1
		if request("action")="updat" then
			call update()
			response.write body
		elseif request("action")="del" then
			call del()
			response.write body
        elseif request("action")="pasword" then
			call pasword()
        elseif request("action")="newpass" then
			call newpass()
			response.write body
		elseif request("action")="add" then
			call addadmin()
		elseif request("action")="savenew" then
			call savenew()
			response.write body
		else
			call userlist()
		end if
		conn.close
		set conn=nothing
	end if

	sub userlist()
%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
                <tr> 
                  <th height=22 colspan=5>管理员管理(点击用户名进行操作)</th>
                </tr>
                <tr align=center> 
                  <td width="30%" height=22 class="forumHeaderBackgroundAlternate"><B>用户名</B></td><td width="25%" class="forumHeaderBackgroundAlternate"><B>上次登陆时间</B></td><td width="15%" class="forumHeaderBackgroundAlternate"><B>上次登陆IP</B></td><td width="15%" class="forumHeaderBackgroundAlternate"><B>操作</B></td>
                </tr>
<%
	set rs=conn.execute("select * from admin order by LastLogin desc")
	do while not rs.eof
%>
                <tr> 
                  <td><a href="admin_admin.asp?id=<%=rs("id")%>&action=pasword"><font color=<%=Forum_body(7)%>><%=rs("username")%></font></a></td><td><font color=<%=Forum_body(7)%>><%=rs("LastLogin")%></font></td><td><font color=<%=Forum_body(7)%>><%=rs("LastLoginIP")%></font></td><td><a href="admin_admin.asp?action=del&id=<%=rs("id")%>&name=<%=rs("username")%>">删除</a></td>
                </tr>
<%
	rs.movenext
	loop
	rs.close
	set rs=nothing
%>
	       </table>
<%
	end sub

	sub del()
	conn.execute("delete from admin where id="&request("id"))
	conn.execute("update [user] set usergroupid=4 where username='"&replace(request("name"),"'","")&"'")
	body="<li>管理员删除成功。"
	end sub

sub pasword()
	set rs=conn.execute("select * from admin where id="&request("id"))
	oldpassword=rs("password")
	oldadduser=rs("adduser")
  %> 
<form action="?action=newpass" method=post>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
               <tr> 
                  <th colspan=2 height=23>管理员资料管理－－密码修改
                  </th>
                </tr>
               <tr > 
            <td width="26%" align="right" class=forumrow>后台登陆名称：</td>
            <td width="74%" class=forumrow>
              <input type=hidden name="oldusername" value="<%=rs("username")%>">
              <input type=text name="username2" value="<%=rs("username")%>">  (可与注册名不同)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow>后台登陆密码：</td>
            <td width="74%" class=forumrow>
              <input type="password" name="password2" value="<%=oldpassword%>">  (可与注册密码不同,如要修改请直接输入)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow height=23>前台用户名称：</td>
            <td width="74%" class=forumrow><%=oldadduser%>
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" class=forumrow> 

              <input type=hidden name="oldadduser" value="<%=oldadduser%>">
              <input type=hidden name="adduser" value="<%=oldadduser%>">
              <input type=hidden name=id value="<%=request("id")%>">
              <input type="submit" name="Submit" value="更 新">
            </td>
          </tr>
        </table>
        </form>

<%       rs.close
         set rs=nothing
end sub

sub newpass()
	dim passnw,usernw,aduser
	set rs=conn.execute("select * from admin where id="&request("id"))
	oldpassword=rs("password")
	if request("username2")="" then
		body="<li>请输入管理员名字。<a href=?>［ <font color=red>返回</font> ］</a>"
		exit sub
	else 
		usernw=trim(request("username2"))
	end if
	if request("password2")="" then
		body="<li>请输入您的密码。<a href=?>［ <font color=red>返回</font> ］</a>"
		exit sub
	elseif trim(request("password2"))=oldpassword then
		passnw=request("password2")
	else
		passnw=md5(request("password2"))
	end if
	if request("adduser")="" then
		body="<li>请输入管理员名字。<a href=?>［ <font color=red>返回</font> ］</a>"
		exit sub
	else 
		aduser=trim(request("adduser"))
	end if

	set rs=server.createobject("adodb.recordset")
	sql="select * from admin where username='"&trim(request("oldusername"))&"'"
	rs.open sql,conn,1,3
	if not rs.eof and not rs.bof then
	rs("username")=usernw
	rs("adduser")=aduser
	rs("password")=passnw
	body="<li>管理员资料更新成功，请记住更新信息。<br> 管理员："&request("username2")&" <BR> 密   码："&request("password2")&" <a href=?>［ <font color=red>返回</font> ］</a>"
	rs.update
	
	end if
	rs.close
	set rs=nothing
end sub


sub addadmin()
%> 
<form action="?action=savenew" method=post>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
               <tr> 
                  <th colspan=2 height=23>管理员管理－－添加管理员
                  </th>
                </tr>
               <tr > 
            <td width="26%" align="right" class=forumrow>后台登陆名称：</td>
            <td width="74%" class=forumrow>
              <input type=text name="username2">  (可与注册名不同)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow>后台登陆密码：</td>
            <td width="74%" class=forumrow>
              <input type="password" name="password2">  (可与注册密码不同)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow height=23>前台用户名称：</td>
            <td width="74%" class=forumrow><input type=text name="username1">  (本选项填写后不允许修改)
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" class=forumrow> 
              <input type="submit" name="Submit" value="添 加">
            </td>
          </tr>
        </table>
        </form>

<%
end sub

sub savenew()
	if request.form("username2")="" then
	body="请输入后台登陆用户名！"
	exit sub
	end if
	if request.form("username1")="" then
	body="请输入前台登陆用户名！"
	exit sub
	end if
	if request.form("password2")="" then
	body="请输入后台登陆密码！"
	exit sub
	end if
	set rs=conn.execute("select username from [user] where username='"&replace(request.form("username1"),"'","")&"'")
	if rs.eof and rs.bof then
	body="您输入的用户名不是一个有效的注册用户！"
	exit sub
	end if
	set rs=conn.execute("select username from admin where username='"&replace(request.form("username2"),"'","")&"'")
	if not (rs.eof and rs.bof) then
	body="您输入的用户名已经在管理用户中存在！"
	exit sub
	end if
	conn.execute("update [user] set usergroupid=1 where username='"&replace(request.form("username1"),"'","")&"'")
	conn.execute("insert into admin (username,[password],adduser) values ('"&replace(request.form("username2"),"'","")&"','"&md5(replace(request.form("password2"),"'",""))&"','"&replace(request.form("username1"),"'","")&"')")
	body="添加成功，请记住新管理员后台登陆信息，如需修改请返回管理员管理！"
end sub
%>