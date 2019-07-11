<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/email.asp"-->
<!--#include file="md5.asp"-->
<%
dim topic,mailbody,sendmsg,SendMail,useremail
dim username,password,repassword
dim answer
stats="密码遗忘"
call nav()
call head_var(1)
if founderr then
	call dvbbs_error()
else
	if request("action")="step1" then
		call step1()
	elseif request("action")="step2" then
		call step2()
	elseif request("action")="step3" then
		call step3()
	else
		call main()
	end if
	if founderr then call dvbbs_error()
end if
call footer()

sub step1()
if request("username")="" then
	Founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请输入您的用户名。"
	exit sub
else
	username=replace(request("username"),"'","")
end if
if Forum_Setting(2)>0 then
	set rs=conn.execute("Select Quesion,Answer,Username,Usergroupid from [user] where username='"&username&"'")
else
	set rs=conn.execute("Select Quesion,Answer,Username,Usergroupid from [user] where username='"&username&"' and UserGroupID>3")
end if
if rs.eof and rs.bof then
	Founderr=true
	errmsg=Errmsg+"<br>"+"<li>您输入的用户名并不存在，请重新输入。<li>或者由于该系统不支持邮件发送，而您的等级为版主以上级别，只能通过联系管理员获得密码。"
elseif rs(3) < 4 then
	Founderr=true
	errmsg=Errmsg+"<br>"+"<li>您是管理员或者是版主，只能通过联系管理员获得密码。"
else
	if rs(0)="" or isnull(rs(0)) then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>该用户没有填写密码问题及答案，只有填写的用户方能继续。"
	else
%>
<form action="lostpass.asp?action=step2" method="post"> 
<table cellpadding=1 cellspacing=1 align=center class=tableborder1>
    <tr>
    <th valign=middle colspan=2 align=center height=25>取回密码（第二步：回答问题）</th></tr>
    <tr>
    <td valign=middle class=tablebody1><b>问题：</b></td>
    <td class=tablebody1 valign=middle><%=rs(0)%></td>
    </tr>
    <tr>
    <td class=tablebody1 valign=middle><b>答案：</b></td>
    <td class=tablebody1 valign=middle><INPUT name="answer" type=text></td>
    </tr>
    <tr>
    <td class=tablebody1 colspan=2>
    <b>说明：</b>请填写您正确的问题答案。</td></tr>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="提 交"></td></tr>
</table>
<input type=hidden value="<%=username%>" name=username>
</form>
<%
		end if
	end if
	rs.close
	set rs=nothing
end sub

sub step2()
	dim answer
	if request("username")="" then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请输入您的用户名。"
		exit sub
	else
		username=replace(request("username"),"'","")
	end if
	if chkpost=false then
   		ErrMsg=ErrMsg+"<Br>"+"<li>您提交的数据不合法，请不要从外部提交发言。"
   		FoundErr=True
		exit sub
	end if
	if request("answer")="" then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请输入您的问题答案。"
		exit sub
	else
		answer=md5(request("answer"))
	end if
	set rs=conn.execute("select answer,quesion from [user] where username='"&username&"' and answer='"&answer&"'")
	if rs.eof and rs.bof then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>您输入的问题答案不正确，请重新输入。"
	else
%>
<form action="lostpass.asp?action=step3" method="post"> 
<table cellpadding=3 cellspacing=1 class=tableborder1 align=center>
    <tr>
    <th valign=middle colspan=2 height=25>取回密码（第三步：修改密码）</th></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>问题：</b></td>
    <td class=tablebody1 valign=middle><%=rs(1)%></td>
    </tr>
    <tr>
    <td class=tablebody1 valign=middle><b>答案：</b></td>
    <td class=tablebody1 valign=middle><%=request("answer")%></td>
    </tr>
    <tr>
    <td class=tablebody1 valign=middle><b>新密码：</b></td>
    <td class=tablebody1 valign=middle><input type=password name=password></td>
    </tr>
    <tr>
    <td class=tablebody1 valign=middle><b>确认新密码：</b></td>
    <td class=tablebody1 valign=middle><input type=password name=repassword></td>
    </tr>
    <tr>
    <td class=tablebody1 colspan=2>
    <b>说明：</b>
	<%if Forum_Setting(2)>0 then%>
	系统将会自动发一封邮件到您注册时填写的邮箱，在打开邮件中的密码激活地址后，您的新密码将正式启用。
	<%else%>
	请填写您的论坛新密码，并记住您所填写信息。<%end if%>
	</td></tr>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="提 交"></td></tr>
</table>
<input type=hidden value="<%=request("answer")%>" name=answer>
<input type=hidden value="<%=username%>" name=username>
</form>
<%
	end if
	rs.close
	set rs=nothing
end sub

sub step3()
	if request("username")="" then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请输入您的用户名。"
		exit sub
	else
		username=replace(request("username"),"'","")
	end if
	if request("answer")="" then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请输入您的问题答案。"
		exit sub
	else
		answer=md5(request("answer"))
	end if
	if request("password")="" or Len(request("password"))>10 or len(request("password"))<6 then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请输入您的新密码(长度不能大于10小于6)。"
		exit sub
	elseif request("repassword")="" then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请再次输入您的新密码。"
		exit sub
	elseif request("password")<>request("repassword") then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>您输入的新密码和确认不一样，请确认您填写的信息。"
		exit sub
	else
		password=md5(request("password"))
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select userpassword,useremail,quesion,userclass,UserGroupID from [user] where username='"&username&"' and answer='"&answer&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>您输入的问题答案不正确，请重新输入。"
	else
		if Forum_Setting(2)>0 then
			repassword=request.form("password")
			answer=request.form("answer")
			password=rs("userpassword")
			useremail=rs("useremail")
			call sendusermail()
			if SendMail="OK" then
			sendmsg="系统已经发送一封邮件到您注册时填写的邮箱，在打开邮件中的密码激活地址后，您的新密码将正式启用。"
			elseif rs("UserGroupID")<4 then
			sendmsg="由于系统错误，给你发送的密码资料未成功，您的等级为版主以上级别，为了安全起见，请联系管理员获得密码。"
			else
			rs("userpassword")=md5(repassword)
			rs.update
			sendmsg="由于系统错误，给您发送的密码资料未成功。您已经修改密码成功，请使用新密码登陆系统。"
			end if
		else
			rs("userpassword")=password
			rs.update
		end if
%>
<form action="login.asp" method="post"> 
<table cellpadding=3 cellspacing=1 class=tableborder1 align=center>
    <tr>
    <th valign=middle colspan=2>取回密码（第四步：修改成功）</th></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>问题：</b></td>
    <td class=tablebody1 valign=middle><%=rs(2)%></td>
    </tr>
    <tr>
    <td class=tablebody1 valign=middle><b>答案：</b></td>
    <td class=tablebody1 valign=middle><%=request("answer")%></td>
    </tr>
    <tr>
    <td class=tablebody1 valign=middle><b>新密码：</b></td>
    <td class=tablebody1 valign=middle><%=request("password")%></td>
    </tr>
    <tr>
    <td class=tablebody1 colspan=2>
    <b>说明：</b>
	<%if Forum_Setting(2)>0 then%>
<%=sendmsg%>
	<%else%>
	请记住您的新密码并使用新密码<a href=login.asp>登陆论坛</a>。
	<%end if%></td></tr>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="返 回"></td></tr>
</table>
</form>
<%
	end if
	rs.close
	set rs=nothing
end sub

sub main()
%>
<form action="lostpass.asp?action=step1" method="post"> 
<table cellpadding=6 cellspacing=1 align=center class=tableborder1>
    <tr>
    <th valign=middle colspan=2>取回密码</b>（第一步：用户名）</th></tr>
    <tr>
    <td class=tablebody1 valign=middle>请输入您的用户名</td>
    <td class=tablebody1 valign=middle><INPUT name="username" type=text></td>
    </tr>
    <tr>
    <td class=tablebody1 colspan=2>
    <b>说明：</b>本操作只能修改您的密码，不能对原密码进行修改，请确认您已经填写了密码问题及答案。</td></tr>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="提 交"></td></tr>
</table>
</form>
<%
end sub

sub sendusermail()
	on error resume next
	topic="您在" & Forum_info(0) & "的密码信息"
	mailbody=mailbody &"<style>A:visited {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline	}"
	mailbody=mailbody &"A:link 	  {	text-decoration: none;}"
	mailbody=mailbody &"A:visited {	text-decoration: none;}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none;}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline}"
	mailbody=mailbody &"BODY   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt;}"
	mailbody=mailbody &"TD	   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt	}</style>"
	mailbody=mailbody &"<TABLE border=0 width='95%' align=center><TBODY><TR>"
	mailbody=mailbody &"<TD valign=middle align=top>"
	mailbody=mailbody &""&htmlencode(username)&"，您好：<br><br>"
	mailbody=mailbody &"欢迎您使用本论坛的密码遗忘功能，<b>假如您不希望更改您的密码，请不要点击下面的激活连接</b>。<br>"
	mailbody=mailbody &"下面是您的密码激活连接，直接将下面的连接地址拷贝到您的浏览器地址并打开就可以将您的新密码激活：<br>"
	mailbody=mailbody &"<a href=http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"lostpass.asp","")&"activepass.asp?username="&htmlencode(username)&"&pass="&password&"&repass="&repassword&"&answer="&answer&">http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"lostpass.asp","")&"activepass.asp?username="&htmlencode(username)&"&pass="&password&"&repass="&repassword&"&answer="&answer&"</a>"
	mailbody=mailbody &"<br><br>"
	mailbody=mailbody &"<center><font color=red>再次感谢您使用本系统，让我们一起来建设这个网上家园！</font>"
	mailbody=mailbody &"</TD></TR></TBODY></TABLE><br><hr width=95% size=1>"
	mailbody=mailbody & Copyright & " &nbsp;&nbsp; " & Version

	select case Forum_Setting(2)
	case 0
	sendmsg="由于系统错误，给您发送的密码资料未成功。请点击右边的连接将您的密码激活：<a href=http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"lostpass.asp","")&"activepass.asp?username="&htmlencode(username)&"&pass="&password&"&repass="&repassword&"><B>激活密码</B></a>"
	case 1
	call jmail(useremail,topic,mailbody)
	case 2
	call Cdonts(useremail,topic,mailbody)
	case 3
	call aspemail(useremail,topic,mailbody)
	case else
	sendmsg="由于系统错误，给您发送的密码资料未成功。请点击右边的连接将您的密码激活：<a href=http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"lostpass.asp","")&"activepass.asp?username="&htmlencode(username)&"&pass="&password&"&repass="&repassword&"><B>激活密码</B></a>"
	end select
end sub

%>