<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="admin_config.asp" -->
<!--#include file="md5.asp"-->
<head>
<title><%=Forum_info(0)%>--管理页面</title>
<%=Eval(""""&gettemplate("Forum_css",4)&"""")%>
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
</head>
<BODY  <%=Forum_body(11)%>>
<%
	if not master or instr(session("flag"),"43")=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		dim body,username2,password2,oldpassword,oldusername
		call main()
		conn.close
		set conn=nothing
	end if

	sub main()
%>
<table cellpadding=0 cellspacing=0 border=0 width=<%=Forum_body(12)%> bgcolor=<%=Forum_body(0)%> align=center>
  <tr>
    <td>
      <table cellpadding=3 cellspacing=1 border=0 width=100%>
        <tr bgcolor='<%=Forum_body(2)%>'>
        <td><font color=<%=Forum_body(6)%>>欢迎<b><%=membername%></b>进入管理中心</font>
        </td>
        </tr>
            <tr bgcolor=<%=Forum_body(4)%>>
              <td width="100%" valign=top><font color=<%=Forum_body(7)%>><p>
<%
	if  request("action")="pasword" then
		call pasword()

        elseif request("action")="newpass" then         
		call newpass()
		response.write body
	else
		call userlist()
	end if
%></font>
		</td>
            </tr>
        </table>
        </td>
    </tr>
</table>
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
<%
	end sub

	sub userlist()
%>
              <table width="80%" border="0" cellspacing="3" cellpadding="1" align=center>
                <tr> 
                  <td colspan=4 bgcolor="<%=Forum_body(3)%>"　align=center> 
                    <font color=<%=Forum_body(7)%>><a href=?><b>>>管理员权限管理<<</b></a>：<br>
                      注意：请选择相应的管理员进行管理操作。(点击相应用户名进行密码修改)</p></font>
                  </td>
                </tr>
                <tr bgcolor="<%=Forum_body(2)%>"> 
                  <td width="35%" height=22><font color=<%=Forum_body(6)%>>用户名</font></td>
                  <td width="30%"><font color=<%=Forum_body(6)%>>上次登陆时间</font></td>
                  <td width="15%"><font color=<%=Forum_body(6)%>>上次登陆IP</font></td>
                  <td width="20%"><font color=<%=Forum_body(6)%>>操作</font></td>
                   
                </tr>
<%
	set rs=conn.execute("select * from admin where adduser='"&membername&"'  order by LastLogin desc")
	do while not rs.eof
%>
                <tr> 
<td><font color=<%=Forum_body(7)%>><%=rs("username")%></font></td>
<td><font color=<%=Forum_body(7)%>><%=rs("LastLogin")%></font></td>
<td><font color=<%=Forum_body(7)%>><%=rs("LastLoginIP")%></font></td>
<td><% if instr(session("flag"),"43")<>0  then %><a href="?id=<%=rs("id")%>&action=pasword">修改</a><%else%>暂停<%end if%></td>
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


sub pasword()

	  set rs=conn.execute("select * from admin where id="&request("id"))
         oldpassword=rs("password")
  %> 
<form action="?action=newpass" method=post>
        <table width="95%" border="0" cellspacing="5" cellpadding="1" align=center>
               <tr> 
                  <td bgcolor="<%=Forum_body(3)%>" align=center colspan=2> 
                    <font color=<%=Forum_body(7)%>><p><a href=?><b>>><b> 管理员资料管理 ：<< <font color=<%=Forum_body(8)%>>密码修改</font> </b></a><br><br>
                      注意：请记下修改完的用户名或密码。</font></p>
                  </td>
                </tr>
               <tr > 
            <td width="26%" align="right" bgcolor="<%=Forum_body(2)%>">后台登陆名称：</td>
            <td width="74%">
              <input type=hidden name="oldusername" value="<%=rs("username")%>">
              <input type=text name="username2" value="<%=rs("username")%>">  (可与注册名不同)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" bgcolor="<%=Forum_body(2)%>">后台登陆密码：</td>
            <td width="74%">
              <input type="password" name="password2" value="<%=oldpassword%>">  (可与注册密码不同,尽量不要过于简单)
            </td>
          </tr>
          <tr bgcolor="<%=Forum_body(3)%>" align="center"> 
            <td colspan="2"> 
              <input type=hidden name=id value="<%=request("id")%>">
              <input type="submit" name="Submit" value="更新">
            </td>
          </tr>
        </table>
        </form>

<%       rs.close
         set rs=nothing
end sub

	sub newpass()
         dim passnw,usernw
         set rs=conn.execute("select * from admin where id="&request("id"))
         oldpassword=rs("password")
               if request("username2")="" then
                   body="<li>请输入管理员名字。<a href=?>［ <font color="&Forum_body(8)&">返回</font> ］</a>"
			exit sub
                  else 
                   usernw=trim(request("username2"))
               end if
                if request("password2")="" then
		  	body="<li>请输入您的密码。<a href=?>［ <font color="&Forum_body(8)&">返回</font> ］</a>"
			exit sub
		elseif trim(request("password2"))=oldpassword then
			passnw=request("password2")
		else
			passnw=md5(request("password2"))
		end if

        set rs=server.createobject("adodb.recordset")
	sql="select * from admin where username='"&trim(request("oldusername"))&"'"
	rs.open sql,conn,1,3
	if not rs.eof and not rs.bof then
	rs("username")=usernw
        rs("password")=passnw
	body="<li>管理员资料更新成功，请记住更新信息。<br> 管理员："&request("username2")&" <BR> 密   码："&request("password2")&" <a href=?>［ <font color="&Forum_body(8)&">返回</font> ］</a>"
	rs.update
	
	end if
	rs.close
	set rs=nothing
 end sub
%>

