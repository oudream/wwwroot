<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!-- #include file="inc/chkinput.asp" -->
<!--#include file="md5.asp"-->
<%
dim username
dim password
dim repassword
dim answer
stats="取回密码"
if request("username")="" or request("pass")="" or request("repass")="" or request("answer")="" then
	Errmsg=Errmsg + "<br><li>非法参数！"
	call nav()
	call head_var(1)
	call dvbbs_error()
else
	username=checkStr(request("username"))
	password=checkStr(request("pass"))
	repassword=md5(checkStr(request("repass")))
	answer=md5(request("answer"))
	sql="select userpassword,userclass,UserGroupID from [user] where username='"&username&"' and userpassword='"&password&"' and answer='"&answer&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		Errmsg=Errmsg + "<li>您返回的用户资料不正确，非法操作！"
		call nav()
		call head_var(1)
		call dvbbs_error()
	else
		if rs("usergroupid")<4 then
		Errmsg=Errmsg + "<li>用户等级为版主或总版主必须和管理员联系取回密码，非法操作！"
		call nav()
		call head_var(1)
		call dvbbs_error()
		else
		rs("userpassword")=repassword
		rs.update
		sucmsg="您的密码成功激活，<a href=login.asp>请登陆</a>"
		call nav()
		call head_var(1)
		call dvbbs_suc()
		end if
	end if
	rs.close
	set rs=nothing
end if
%>