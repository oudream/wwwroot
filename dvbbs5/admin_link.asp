<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content=""Microsoft FrontPage 3.0"" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		dim body
		dim readme,Tlink
		call main()
		set rs=nothing
		conn.close
		set conn=nothing
	end if

sub main()
	if request("action") = "add" then 
		call addlink()
	elseif request("action")="edit" then
		call editlink()
	elseif request("action")="savenew" then
		call savenew()
	elseif request("action")="savedit" then
		call savedit()
	elseif request("action")="del" then
		call del()
	elseif request("action")="orders" then
		call orders()
	elseif request("action")="updatorders" then
		call updateorders()
	else
		call linkinfo()
	end if
	response.write body
end sub

sub addlink()
%>
<form action="admin_link.asp?action=savenew" method = post>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
  <tr> 
    <th width="100%" colspan=2 height=25>添加联盟论坛 </th>
  </tr>
  <tr> 
    <td width="40%" height=25 class=forumrow>论坛名称 </td>
    <td width="60%" height=25 class=forumrow> 
      <input type="text" name="name" size=40>
    </td>
  </tr>
  <tr> 
    <td width="40%" height=25 class=forumrow>连接URL </td>
    <td width="60%" class=forumrow> 
      <input type="text" name="url" size=40>
    </td>
  </tr>
  <tr> 
    <td width="40%" height=25 class=forumrow>连接LOGO地址 </td>
    <td width="60%" class=forumrow> 
      <input type="text" name="logo" size=40>
    </td>
  </tr>
  <tr> 
    <td width="40%" height=25 class=forumrow>论坛简介 </td>
    <td width="60%" class=forumrow> 
      <input type="text" name="readme" size=40>
    </td>
  </tr>
  <tr> 
    <td width="40%" height=25 class=forumrow>在首页是文字连接还是LOGO连接 </td>
    <td width="60%" class=forumrow> 
      文字连接<input type="radio" name="islogo" value=0 checked>&nbsp;&nbsp;LOGO连接<input type="radio" name="islogo" value=1>
    </td>
  </tr>
  <tr> 
    <td height="15" colspan="2" class=forumrow> 
        <input type="submit" name="Submit" value="添 加">
    </td>
  </tr>
</table>
</form>
<%
end sub

sub editlink()
	set rs= server.createobject ("adodb.recordset")
	sql = "select * from bbslink where id="&Request("id")
	rs.open sql,conn,1,1
%>
<form action="admin_link.asp?action=savedit" method=post>
<input type=hidden name=id value=<%=Request("id")%>>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
  <tr> 
    <th width="100%" colspan=2 height=25>编辑联盟论坛</th>
  </tr>
  <tr> 
    <td width="40%" class=forumrow><font color="<%=Forum_body(7)%>">论坛名称： </font></td>
    <td width="60%" class=forumrow> 
      <input type="text" name="name" size=40 value=<%=rs("boardname")%>>
    </td>
  </tr>
  <tr> 
    <td width="40%" class=forumrow><font color="<%=Forum_body(7)%>">连接URL：</font> </td>
    <td width="60%" class=forumrow> 
      <input type="text" name="url" size=40 value=<%=rs("url")%>>
    </td>
  </tr>
  <tr> 
    <td width="40%" class=forumrow><font color="<%=Forum_body(7)%>">连接LOGO地址： </font></td>
    <td width="60%" class=forumrow> 
      <input type="text" name="logo" size=40 value="<%=rs("logo")%>">
    </td>
  </tr>
  <tr> 
    <td width="40%" class=forumrow><font color="<%=Forum_body(7)%>">论坛简介： </font></td>
    <td width="60%" class=forumrow> 
      <input type="text" name="readme" size=40 value=<%=rs("readme")%>>
    </td>
  </tr>
  <tr> 
    <td width="40%" height=25 class=forumrow>在首页是文字连接还是LOGO连接 </td>
    <td width="60%" class=forumrow> 
      文字连接<input type="radio" name="islogo" value=0 <%if rs("islogo")=0 then%>checked<%end if%>>&nbsp;&nbsp;LOGO连接<input type="radio" name="islogo" value=1 <%if rs("islogo")=1 then%>checked<%end if%>>
    </td>
  </tr>
  <tr> 
    <td height="15" colspan="2" class=forumrow> 
      <div align="center"> 
        <input type="submit" name="Submit" value="修 改">
      </div>
    </td>
  </tr>
</table>
</form>
<%
	rs.close
	set rs=nothing
end sub

sub linkinfo()
	set rs= server.createobject ("adodb.recordset")
	sql = " select * from bbslink order by id"
	rs.open sql,conn,1,1       
%> 
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
              <tr> 
                <th height="22" colspan=4>联盟论坛列表 | <a href="admin_link.asp?action=add"><font color=#FFFFFF>增加新的联盟论坛</font></a></th>
              </tr>
			<tr align=center>
			<td width="10%" height=25 class="forumHeaderBackgroundAlternate"><B>序号</B></td>
			<td width="40%" class="forumHeaderBackgroundAlternate"><B>名称</B></td>
			<td width="20%" class="forumHeaderBackgroundAlternate"><B>是否LOGO</B></td>
			<td width="30%" class="forumHeaderBackgroundAlternate"><B>编辑</B></td>
			</tr>
<%
	do while not rs.eof
         Tlink=split(rs(2),"$")
%>
			<tr align=center>
			<td height=25 class=forumrow><font color=red><%=rs("id")%></font></td>
			<td class=forumrow><a href=<%=rs("url")%> target=_blank><%=rs("boardname")%></a></td>
			<td class=forumrow><%if rs("islogo")=1 then%>是<%else%>否<%end if%></td>
			<td class=forumrow><a href="admin_link.asp?action=orders&id=<%=rs("id")%>">排序</a>  <a href="admin_link.asp?action=edit&id=<%=rs("id")%>">编辑</a>  <a href="admin_link.asp?action=del&id=<%=rs("id")%>">删除</a></td>
			</tr>
<%
	rs.movenext
	loop
	rs.Close
	set rs=nothing
%>
            </table>
<%
end sub

sub savenew()
if Request("url")<>"" and Request("readme")<>"" and request("name")<>"" then
	dim linknum
	set rs= server.createobject ("adodb.recordset")
	sql = "select * from bbslink order by id desc"
	rs.Open sql,conn,1,3
	if rs.eof and rs.bof then
	linknum=1
	else
	linknum=rs("id")+1
	end if
	rs.AddNew 
	rs("id")=linknum
	rs("boardname") = Trim(Request.Form ("name"))
	rs("readme") =  Trim(Request.Form ("readme"))
	rs("logo") = trim(request.Form("logo"))
	rs("url") = Request.Form ("url")
	rs("islogo")=request.Form("islogo")
	rs.Update 
	rs.Close
	set rs=nothing
	body=body+"<br>"+"更新成功，请继续其他操作。"
else
	body=body+"<br>"+"请输入完整联盟论坛信息。"
end if
end sub

sub savedit()
	set rs= server.createobject ("adodb.recordset")
	sql = "select * from bbslink where id="&request("id")
	rs.Open sql,conn,1,3
	if rs.eof and rs.bof then
	body=body+"<br>"+"错误，没有找到联盟论坛。"
	else
	rs("boardname") = Trim(Request.Form ("name"))
	rs("readme") =  Trim(Request.Form ("readme"))
	rs("logo")=trim(request.Form("logo"))
	rs("url") = Request.Form ("url")
	rs("islogo")=request.Form("islogo")
	rs.Update
	end if 
	rs.Close
	set rs=nothing
	body=body+"<br>"+"更新成功，请继续其他操作。"
end sub

sub del
	dim id
	id = request("id")
	sql="delete from bbslink where id="+id
	conn.Execute(sql)
	body=body+"<br>"+"删除成功，请继续其他操作。"
end sub

sub orders()
%><br>
            <table width="95%" border="0" cellspacing="3" cellpadding="0" align=center>
              <tr> 
                <td height="22"><font color="<%=Forum_body(7)%>"><b>联盟论坛重新排序</b><br>
注意：请在相应论坛的排序表单内输入相应的排列序号，<font color=red>注意不能和别的联盟论坛有相同的排列序号</font>。</font>
		</td>
              </tr>
	      <tr>
              <tr bgcolor="<%=Forum_body(3)%>"> 
                <td height="22" align=center><a href="admin_link.asp?action=add"><font color="<%=Forum_body(7)%>">增加新的联盟论坛</font></a></td>
              </tr>
	      <tr>
		<td><font color="<%=Forum_body(7)%>">
<%
	set rs= server.createobject ("adodb.recordset")
	sql="select * from bbslink where id="&cstr(request("id"))
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "没有找到相应的联盟论坛。"
	else
		response.write "<form action=admin_link.asp?action=updatorders method=post>"
		response.write ""&rs("boardname")&"  <input type=text name=newid size=2 value="&rs("id")&">"
		response.write "<input type=hidden name=id value="&request("id")&">"
		response.write "<input type=submit name=Submit value=修改></form>"
	end if
	rs.close
	set rs=nothing
%></font>
		</td>
	      </tr>
              <tr bgcolor="<%=Forum_body(3)%>"> 
                <td height="22" align=center><a href="admin_link.asp?action=add"><font color="<%=Forum_body(7)%>">增加新的联盟论坛</font></a></td>
              </tr>
            </table>
<%
end sub

sub updateorders()
if isnumeric(request("id")) and isnumeric(request("newid")) and request("newid")<>request("id") then
	set rs=conn.execute("select id from bbslink where id="&request("newid"))
	if rs.eof and rs.bof then
	sql="update bbslink set id="&request("newid")&" where id="&cstr(request("id"))
	conn.execute(sql)
	response.write "更新成功！"
	else
	response.write "更新失败，您指定了和其他联盟论坛相同的序号！"
	end if
else
	response.write "更新失败！您输入的字符不合法，或者输入了和原来相同的序号！"
end if
end sub
%>