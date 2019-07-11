<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!-- #include file="md5.asp" -->
<%
	stats="论坛管理登陆"
	dim username
	dim password
	dim ip
	templateused=templateused & "Forum_adminlogin"
	cachetemplate templateused,0
	call head_var(1)
	if not founduser then
		Errmsg="<li>您不是系统管理员！"
		call Dvbbs_error(1)
	else
		if request("action")="chklogin" then
		call chklogin()
		if founderr then call Dvbbs_error(1)
		else
		dim num1
		dim rndnum
		Randomize
		Do While Len(rndnum)<4
		num1=CStr(Chr((57-48)*rnd+48))
		rndnum=rndnum&num1
		loop
		session("verifycode")=rndnum
		endtime=timer()
		response.write eval(""""&gettemplate("Forum_adminlogin",126)&"""")
		end if
	end if
	call activeonline()
	call footer()
	sub chklogin()
	username=trim(replace(request("username"),"'",""))
	password=md5(trim(replace(request("password"),"'","")))
	if request("verifycode")="" then
		errmsg=errmsg+"<br>"+"<li>请返回输入确认码。"
		founderr=true
	elseif session("verifycode")="" then
		errmsg=errmsg+"<br>"+"<li>请不要重复提交，如需重新登陆请返回登陆页面。"
		founderr=true
	elseif session("verifycode")<>trim(request("verifycode")) then
		errmsg=errmsg+"<br>"+"<li>您输入的确认码和系统产生的不一致，请重新输入。"
		founderr=true
	end if
	session("verifycode")=""
	if username="" or password="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请输入您的用户名或密码。"
	end if
	if founderr then exit sub
	ip=Request.ServerVariables("REMOTE_ADDR")
	set rs=conn.execute("select * from admin where username='"&username&"' and adduser='"&membername&"'")
	if rs.eof and rs.bof then
		rs.close
		set rs=nothing
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>您输入的用户名和密码不正确或者您不是系统管理员。<br><li>请<a href=admin_login.asp>重新输入</a>您的密码。"
		exit sub
	else
		if trim(rs("password"))<>password then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>您输入的用户名和密码不正确或者您不是系统管理员。<br><li>请<a href=admin_login.asp>重新输入</a>您的密码。"
		exit sub
		else
		if isnull(rs("flag")) or rs("flag")="" then
		rs.close
		set rs=nothing
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>您没有权限进入管理页面。"
		exit sub
		else
		session("flag")=rs("flag")
		session.timeout=45
		conn.execute("update admin set LastLogin=Now(),LastLoginIP='"&ip&"' where username='"&username&"'")
		rs.close
		set rs=nothing
		response.write "<script>location.href='admin_main.asp'</script>"
		end if
		end if
	end if
	end sub
%>