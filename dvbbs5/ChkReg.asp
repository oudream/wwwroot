<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/chkinput.asp"-->
<%
	stats="注册用户"
	dim username
	dim useremail
	if request("username")="" or strLength(request("username"))>20 then
		errmsg=errmsg+"<br>"+"<li>请输入您的用户名(长度不能大于20)。"
		founderr=true
	else
		username=trim(request("username"))
	end if
	if Instr(request("username"),"=")>0 or Instr(request("username"),"%")>0 or Instr(request("username"),chr(32))>0 or Instr(request("username"),"?")>0 or Instr(request("username"),"&")>0 or Instr(request("username"),";")>0 or Instr(request("username"),",")>0 or Instr(request("username"),"'")>0 or Instr(request("username"),",")>0 or Instr(request("username"),chr(34))>0 or Instr(request("username"),chr(9))>0 or Instr(request("username"),"")>0 or Instr(request("username"),"$")>0 then
	errmsg=errmsg+"<br>"+"<li>用户名中含有非法字符。"
		founderr=true
	else
		username=trim(request("username"))
	end if

	splitwords=split(splitwords,",")
	for i = 0 to ubound(splitwords)
		if instr(username,splitwords(i))>0 then
			errmsg=errmsg+"<br>"+"<li>您输入的用户名包含系统禁止注册字符。"
			founderr=true
			exit for
		end if
	next
	
	if IsValidEmail(trim(request("email")))=false then
   		errmsg=errmsg+"<br>"+"<li>您的Email有错误。"
   		founderr=true
	else
		useremail=checkStr((request("email")))
	end if
	if not founderr then
	if cint(Forum_Setting(24))=1 then
	sql="select username,useremail from [user] where username='"&username&"' or useremail='"&useremail&"'"
	else
	sql="select username,useremail from [user] where username='"&username&"'"
	end if
	set rs=conn.execute(sql)
	if not rs.eof and not rs.bof then
		if cint(Forum_Setting(24))=1 and rs("useremail")=useremail then
		errmsg=errmsg+"<br>"+"<li>对不起，本论坛已经限制了一个Email只能注册一个帐号，请重新选择您的Email。"
		else
		errmsg=errmsg+"<br>"+"<li>对不起，您输入的用户名已经被注册。"
		end if
		founderr=true
	else
		founderr=false
	end if
	rs.close
	set rs=nothing
	end if
%>
<html><head>
<meta NAME=GENERATOR Content="Microsoft FrontPage 4.0" CHARSET=GB2312>
<title><%=Forum_info(0)%>--<%=stats%></title>
<!--#include file=inc/Forum_css.asp--></head>
<body <%=Forum_body(11)%>>
<p>&nbsp;</p>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TBODY> 
<TR align=middle> 
<Th height=24>用户检测情况</Th>
</TR>
<TR> 
<Td class=tablebody1 height=24>
<%
if founderr then
response.write Errmsg
else
response.write "恭喜，您所填写的用户和Email通过检测，可以正常注册！<br>请继续将您的注册信息填写完整，谢谢！"
end if
%></TD>
</TR>
</TBODY>
</TABLE>
<%call footer()%>