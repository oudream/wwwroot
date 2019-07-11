<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!--#include file=inc/email.asp-->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		dim topic,mailbody,useremail
		i=1
		set rs= server.createobject ("adodb.recordset")
		if request("action")="send" then
			call sendmail()
		else
			call mail()
		end if
	end if

sub mail()
%>
<form action="admin_mailist.asp?action=send" method=post>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
                <tr> 
                  <th colspan="2">论坛邮件列表
                  </th>
                </tr>
                <tr> 
                  <td width="20%" class=forumrow>注意事项：</td>
		  <td class=forumrow>在完整填写以下表单后点击发送，信息将发送到所有注册时完整填写了信箱的用户，邮件列表的使用将消耗大量的服务器资源，请慎重使用。</td>
                </tr>
                <tr> 
                  <td width="20%" class=forumrow>邮件标题：</td>
		  <td class=forumrow><input type=text name=topic size=60></td>
                </tr>
                <tr> 
                  <td width="20%" class=forumrow>邮件内容：</td>
		  <td class=forumrow><textarea cols=80 rows=6 name="content"></textarea></td>
                </tr>
                <tr>  <td width="20%" class=forumrow></td>
                  <td height=20 class=forumrow>
<input type=Submit value="发 送" name=Submit"> &nbsp; <input type="reset" name="Clear" value="清 除">
                  </td>
                </tr>
              </table>
</form>
<%
end sub

sub sendmail()
	if request("topic")="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入邮件标题。"
		founderr=true
	else
		topic=request("topic")
	end if
	if request("content")="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入邮件内容。"
		founderr=true
	else
		mailbody=request("content")
	end if
	if founderr=false then
	on error resume next
	sql="select username,useremail from [user]"
	rs.open sql,conn,1,1
	if not rs.eof and not rs.bof then
		alluser=rs.recordcount
		do while not rs.eof
		if rs("useremail")<>"" then
		useremail=rs("useremail")
			if Forum_Setting(2)=0 then
				errmsg=errmsg+"<br>"+"<li>本论坛不支持发送邮件。"
				exit sub
			elseif Forum_Setting(2)=1 then
				call jmail(useremail,topic,mailbody)
			elseif Forum_Setting(2)=2 then
				call Cdonts(useremail,topic,mailbody)
			elseif Forum_Setting(2)=3 then
				call aspemail(useremail,topic,mailbody)
			end if
		i=i+1
		end if
		rs.movenext
		loop
		Errmsg=Errmsg+"<br>"+"<li>成功发送"&i&"封邮件。"
	end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	end if
	response.write ""&Errmsg&""
end sub
%>