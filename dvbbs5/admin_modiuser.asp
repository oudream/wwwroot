<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file=inc/char.asp-->
<!--#include file=inc/grade.asp-->
<!--#include file=md5.asp-->
<title><%=Forum_info(0)%>--管理页面</title>
<%=Eval(""""&gettemplate("Forum_css",4)&"""")%>

<BODY <%=Forum_body(11)%>>
<%
	if not master or instr(session("flag"),"11")=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		call main()
		set rs=nothing
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
        <td><font color="<%=Forum_body(6)%>">欢迎<b><%=membername%></b>进入管理页面</font>
        </td>
        </tr>
            <tr bgcolor=<%=Forum_body(4)%>>
              <td width="100%" valign=top><font color="<%=Forum_body(7)%>">
<%
if request("action")="save" then
call update()
else
call userinfo()
end if
%></font>
	      </td>
            </tr>
        </table>
        </td>
    </tr>
</table>
<%
end sub

sub userinfo()
	dim username
	username=trim(replace(request("name"),"'",""))
	set rs=server.createobject("adodb.recordset")
   	sql="select * from [User] where username='"&UserName&"'"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		errmsg=errmsg+"<br>"+"<li>该用户名不存在。"
		call dvbbs_error()
		exit sub
	else
%>
<form method="POST" action=admin_modiuser.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center bordercolor=<%=Forum_body(1)%>>
<tr bgcolor=<%=Forum_body(3)%>> 
<td height="11" colspan="2" ><font color="<%=Forum_body(7)%>"><b><%=htmlencode(rs("username"))%>的个人资料</b></font></td>
</tr>
<tr > 
<td height="11" ><font color="<%=Forum_body(7)%>">用户头衔（管理员设定）</font></td>
<td height="11" > 
<input type="text" name="title" size="35" value="<%=rs("title")%>">
</td>
</tr>
<tr> 
<td width="41%" height="18"><font color="<%=Forum_body(7)%>">用户名</font></td>
<td width="59%" height="18"> 
<input type="text" name="userName" size="35" value="<%=htmlencode(rs("username"))%>">
<input type="hidden" name="Name" value="<%=htmlencode(rs("username"))%>">
</td>
</tr>
<tr> 
<td width="41%" height="18"><font color="<%=Forum_body(7)%>">用户密码</font></td>
<td width="59%" height="18"> 
        <input type="text" name="password" size="35" value="<%=htmlencode(rs("userpassword"))%>">
</td>
</tr>
<tr>
        <td width="40%">用户密码问题</td>   
        <td width="60%">    
          <input type=text name="quesion" size=30 value="<%=rs("quesion")%>">
        </td>   
      </tr>
<tr>
        <td width="40%">问题答案</td>   
        <td width="60%">    
          <input type=text name="answer" size=30 value="<%=htmlencode(rs("answer"))%>">
        </td>   
      </tr>
<tr> 
<td width="41%" height="18"><font color="<%=Forum_body(7)%>">邮件地址</font></td>
<td width="59%" height="18"> 
<input type="text" name="userEmail" size="35" value="<%=rs("userEmail")%>">
</td>
</tr>
<tr> 
<td width="41%" height="18"><font color="<%=Forum_body(7)%>">个人主页</font></td>
<td width="59%" height="18"> 
<input type="text" name="homepage" size="35" value="<%=rs("homepage")%>">
</td>
</tr>
<tr> 
<td width="41%" height="8"><font color="<%=Forum_body(7)%>">发表文章</font></td>
<td width="59%" height="-2"> 
<input type="text" name="article" size="35" value="<%=rs("article")%>">
</td>
</tr>
<tr> 
<td width="41%" height="9"><font color="<%=Forum_body(7)%>">财产总数</font></td>
<td width="59%" height="0"> 
<input type="text" name="userWealth" size="35" value="<%=rs("userWealth")%>">
</td>
</tr>
<tr> 
<td width="41%" height="8"><font color="<%=Forum_body(7)%>">经 验 值</font></td>
<td width="59%" height="8"> 
<input type="text" name="userEP" size="35" value="<%=rs("userEP")%>">
</td>
</tr>
<tr> 
<td width="41%" height="9"><font color="<%=Forum_body(7)%>">魅 力 值</font></td>
<td width="59%" height="9">
<input type="text" name="userCP" size="35" value="<%=rs("userCP")%>">
</td>
</tr>

<tr> 
<td width="41%" height="18">头像地址</td>
<td width="59%" height="18"> 
<input type="text" name="FACE" size="60" value="<%=rs("FACE")%>">
</td>
</tr>
</tr>
 <tr>    
        <td valign=top width="40%"><B>签&nbsp;&nbsp;&nbsp;&nbsp;名</B>：<BR>不能超过 300 个字符  </font></td>   
        <td width="60%">    
          <textarea name="Signature" rows=5 cols=60 wrap=PHYSICAL><%if rs("sign")<>"" then%>
<%
dim signtrue
signtrue=replace(rs("sign"),"<BR>",chr(13))
signtrue=replace(signtrue,"&nbsp;"," ")
%><%=signtrue%><%end if%></textarea>   
        </td>   
      </tr>
<tr> 
<td width="41%" height="18"><font color="<%=Forum_body(7)%>">用户等级</font></td>
<td width="59%" height="18"> 
<select name="userclass">
<%
	dim userclass
	userclass=(rs("userclass"))
%>
<option value="1" <%if grade(userclass)=(grade(1)) then%>selected<%end if%>><%=grade(1)%> 
<option value="2" <%if grade(userclass)=(grade(2)) then%>selected<%end if%>><%=grade(2)%> 
<option value="3" <%if grade(userclass)=(grade(3)) then%>selected<%end if%>><%=grade(3)%> 
<option value="4" <%if grade(userclass)=(grade(4)) then%>selected<%end if%>><%=grade(4)%> 
<option value="5" <%if grade(userclass)=(grade(5)) then%>selected<%end if%>><%=grade(5)%> 
<option value="6" <%if grade(userclass)=(grade(6)) then%>selected<%end if%>><%=grade(6)%> 
<option value="7" <%if grade(userclass)=(grade(7)) then%>selected<%end if%>><%=grade(7)%> 
<option value="8" <%if grade(userclass)=(grade(8)) then%>selected<%end if%>><%=grade(8)%> 
<option value="9" <%if grade(userclass)=(grade(9)) then%>selected<%end if%>><%=grade(9)%> 
<option value="10" <%if grade(userclass)=(grade(10)) then%>selected<%end if%>><%=grade(10)%> 
<option value="11" <%if grade(userclass)=(grade(11)) then%>selected<%end if%>><%=grade(11)%> 
<option value="12" <%if grade(userclass)=(grade(12)) then%>selected<%end if%>><%=grade(12)%> 
<option value="13" <%if grade(userclass)=(grade(13)) then%>selected<%end if%>><%=grade(13)%> 
<option value="14" <%if grade(userclass)=(grade(14)) then%>selected<%end if%>><%=grade(14)%> 
<option value="15" <%if grade(userclass)=(grade(15)) then%>selected<%end if%>><%=grade(15)%> 
<option value="16" <%if grade(userclass)=(grade(16)) then%>selected<%end if%>><%=grade(16)%> 
<option value="17" <%if grade(userclass)=(grade(17)) then%>selected<%end if%>><%=grade(17)%> 
<option value="18" <%if grade(userclass)=(grade(18)) then%>selected<%end if%>><%=grade(18)%> 
<option value="19" <%if grade(userclass)=(grade(19)) then%>selected<%end if%>><%=grade(19)%> 
<%if userclass=20 then%>
<option value="20" <%if grade(userclass)=(grade(20)) then%>selected<%end if%>><%=grade(20)%> 
<%end if%>
</select>
</td>
</tr>
<tr> 
<td width="41%" height="18"><font color="<%=Forum_body(7)%>">锁定用户</font></td>
<td width="59%" height="18"> 
<select name="lockuser">
<option value="0" <%if rs("lockuser")=0 then%>selected<%end if%>>否 
<option value="1" <%if rs("lockuser")=1 then%>selected<%end if%>>是 
</select>
</td>
</tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<td height="23" colspan="2" > 
<input type="submit" name="Submit" value="更 新">
</td>
</tr>
</table>
</form>
<%
	end if
	rs.close
	set rs=nothing
end sub

sub update()
	dim username
	username=trim(replace(request("name"),"'",""))
	set rs=server.createobject("adodb.recordset")
   	sql="select * from [User] where username='"&UserName&"'"
	'response.write sql
	'response.end
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		errmsg=errmsg+"<br>"+"<li>该用户名不存在。"
		call dvbbs_error()
		exit sub
	else
		rs("title")=request.form("title")
		rs("username")=request.form("username")
		if rs("userpassword")<>request.form("password") then
		rs("userpassword")=md5(request.form("password"))
		end if
                rs("quesion")=request.form("quesion")
                if rs("answer")<>request.form("answer") then
		rs("answer")=md5(request.form("answer"))
                end if
                rs("face")=request.form("face")
		rs("useremail")=request.form("useremail")
		rs("homepage")=request.form("homepage")
		rs("article")=request.form("article")
		rs("userclass")=request.form("userclass")
		rs("lockuser")=request.form("lockuser")
		rs("userWealth")=request.form("userWealth")
		rs("userEP")=request.form("userEP")
		rs("userCP")=request.form("userCP")
                rs("sign")=trim(request("Signature"))
		rs.update
		if request.form("username")<>rs("username") then
		conn.execute("update bbs1 set username='"&request.form("username")&"' where username='"&rs("username")&"'")
		conn.execute("update message set sender='"&request.form("username")&"' where sender='"&rs("username")&"'")
		conn.execute("update message set incept='"&request.form("username")&"' where incept='"&rs("username")&"'")
		end if
        rs.close
	end if
	set rs=nothing
%><center><p><b>更新用户数据成功！</b>
<%
end sub
%>