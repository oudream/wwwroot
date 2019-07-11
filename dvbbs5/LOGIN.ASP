<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/char_login.asp"-->
<!--#include file="md5.asp"-->
<%
'=========================================================
' File: login.asp
' Version:5.0
' Date: 2002-9-10
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

stats="论坛登陆"
dim url,comeurl

call nav()
call head_var(1)
if request("action")="chk" then
	call chklogin()
	if founderr then call dvbbs_error()
else
	url=Request.ServerVariables("HTTP_REFERER")
%>
<form action="login.asp?action=chk" method=post> 
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
    <tr>
    <th valign=middle colspan=2 align=center height=25>请输入您的用户名、密码登陆</td></tr>
    <tr>
    <td valign=middle class=tablebody1>请输入您的用户名</td>
    <td valign=middle class=tablebody1><INPUT name=username type=text> &nbsp; <a href=reg.asp>没有注册？</a></td></tr>
    <tr>
    <td valign=middle class=tablebody1>请输入您的密码</font></td>
    <td valign=middle class=tablebody1><INPUT name=password type=password> &nbsp; <a href=lostpass.asp>忘记密码？</a></td></tr>
    <tr>
    <td class=tablebody1 valign=top width=30% ><b>Cookie 选项</b><BR> 请选择你的 Cookie 保存时间，下次访问可以方便输入。</td>
    <td valign=middle class=tablebody1>                <input type=radio name=CookieDate value=0 checked>不保存，关闭浏览器就失效<br>
                <input type=radio name=CookieDate value=1>保存一天<br>
                <input type=radio name=CookieDate value=2>保存一月<br>
                <input type=radio name=CookieDate value=3>保存一年<br>                </td></tr>
    <tr>
    <td valign=top width=30% class=tablebody1><b>隐身登陆</b><BR> 您可以选择隐身登陆，论坛会员将在用户列表看不到您的信息。</td>
    <td valign=middle class=tablebody1>                <input type=radio name=userhidden value=2 checked>正常登陆<br>
                <input type=radio name=userhidden value=1>隐身登陆<br>
                </td></tr>
	<input type=hidden name=comeurl value=<%=url%>>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name=submit value="登 陆"></td></tr></table>
</form>
<%
end if
call footer()

sub chklogin()
	dim UserIP
	dim username
	dim userclass
	dim password
	dim article
	dim cookies_path_s,cookies_path_d,cookies_path,usercookies
	UserIP=Request.ServerVariables("REMOTE_ADDR")
	if request("username")="" then
		errmsg=errmsg+"<br>"+"<li>请输入您的用户名。"
		founderr=true
	else
		username=trim(checkStr(request("username")))
	end if
	if request("password")="" then
		errmsg=errmsg+"<br>"+"<li>请输入您的密码。"
		founderr=true
	else
		password=md5(trim(checkStr(request("password"))))
	end if
	'判断更新cookies目录
	cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
	cookies_path_d=ubound(cookies_path_s)
	cookies_path="/"
	for i=1 to cookies_path_d-1
		if not (cookies_path_s(i)="upload" or cookies_path_s(i)="admin") then 	cookies_path=cookies_path&cookies_path_s(i)&"/"
	next
	if cookiepath<>cookies_path then
		cookiepath=cookies_path
		conn.execute("update config set cookiepath='"&cookiepath&"'")
	end if
	if founderr then exit sub
	
	usercookies=request("CookieDate")
	if chkuserlogin(username,password,usercookies,1)=false then
		errmsg=errmsg+"<br>"+"<li>您的用户名并不存在，或者您的密码错误，或者您的帐号已被管理员锁定。"
		founderr=true
		exit sub
	end if
	dim comeurlname
	if instr(lcase(request("comeurl")),"reg.asp")>0 or instr(lcase(request("comeurl")),"login.asp")>0 or  trim(request("comeurl"))="" then
		comeurlname=""
		comeurl="index.asp"
	else
		comeurl=request("comeurl")
		comeurlname="<li><a href="&request("comeurl")&">"&request("comeurl")&"</a></li>"
	end if
%>
<meta HTTP-EQUIV=REFRESH CONTENT='2; URL=<%=comeurl%>'>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1 >
<tr>
<th height=25>登陆成功：<%=Forum_info(0)%>欢迎您的到来</th>
</tr>
<tr><td class=tablebody1><br>
<ul><%=comeurlname%><li><a href=index.asp>返回首页</a></li></ul>
</td></tr>
</table>
<%
end sub
%>