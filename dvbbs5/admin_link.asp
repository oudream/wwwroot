<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content=""Microsoft FrontPage 3.0"" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
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
    <th width="100%" colspan=2 height=25>���������̳ </th>
  </tr>
  <tr> 
    <td width="40%" height=25 class=forumrow>��̳���� </td>
    <td width="60%" height=25 class=forumrow> 
      <input type="text" name="name" size=40>
    </td>
  </tr>
  <tr> 
    <td width="40%" height=25 class=forumrow>����URL </td>
    <td width="60%" class=forumrow> 
      <input type="text" name="url" size=40>
    </td>
  </tr>
  <tr> 
    <td width="40%" height=25 class=forumrow>����LOGO��ַ </td>
    <td width="60%" class=forumrow> 
      <input type="text" name="logo" size=40>
    </td>
  </tr>
  <tr> 
    <td width="40%" height=25 class=forumrow>��̳��� </td>
    <td width="60%" class=forumrow> 
      <input type="text" name="readme" size=40>
    </td>
  </tr>
  <tr> 
    <td width="40%" height=25 class=forumrow>����ҳ���������ӻ���LOGO���� </td>
    <td width="60%" class=forumrow> 
      ��������<input type="radio" name="islogo" value=0 checked>&nbsp;&nbsp;LOGO����<input type="radio" name="islogo" value=1>
    </td>
  </tr>
  <tr> 
    <td height="15" colspan="2" class=forumrow> 
        <input type="submit" name="Submit" value="�� ��">
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
    <th width="100%" colspan=2 height=25>�༭������̳</th>
  </tr>
  <tr> 
    <td width="40%" class=forumrow><font color="<%=Forum_body(7)%>">��̳���ƣ� </font></td>
    <td width="60%" class=forumrow> 
      <input type="text" name="name" size=40 value=<%=rs("boardname")%>>
    </td>
  </tr>
  <tr> 
    <td width="40%" class=forumrow><font color="<%=Forum_body(7)%>">����URL��</font> </td>
    <td width="60%" class=forumrow> 
      <input type="text" name="url" size=40 value=<%=rs("url")%>>
    </td>
  </tr>
  <tr> 
    <td width="40%" class=forumrow><font color="<%=Forum_body(7)%>">����LOGO��ַ�� </font></td>
    <td width="60%" class=forumrow> 
      <input type="text" name="logo" size=40 value="<%=rs("logo")%>">
    </td>
  </tr>
  <tr> 
    <td width="40%" class=forumrow><font color="<%=Forum_body(7)%>">��̳��飺 </font></td>
    <td width="60%" class=forumrow> 
      <input type="text" name="readme" size=40 value=<%=rs("readme")%>>
    </td>
  </tr>
  <tr> 
    <td width="40%" height=25 class=forumrow>����ҳ���������ӻ���LOGO���� </td>
    <td width="60%" class=forumrow> 
      ��������<input type="radio" name="islogo" value=0 <%if rs("islogo")=0 then%>checked<%end if%>>&nbsp;&nbsp;LOGO����<input type="radio" name="islogo" value=1 <%if rs("islogo")=1 then%>checked<%end if%>>
    </td>
  </tr>
  <tr> 
    <td height="15" colspan="2" class=forumrow> 
      <div align="center"> 
        <input type="submit" name="Submit" value="�� ��">
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
                <th height="22" colspan=4>������̳�б� | <a href="admin_link.asp?action=add"><font color=#FFFFFF>�����µ�������̳</font></a></th>
              </tr>
			<tr align=center>
			<td width="10%" height=25 class="forumHeaderBackgroundAlternate"><B>���</B></td>
			<td width="40%" class="forumHeaderBackgroundAlternate"><B>����</B></td>
			<td width="20%" class="forumHeaderBackgroundAlternate"><B>�Ƿ�LOGO</B></td>
			<td width="30%" class="forumHeaderBackgroundAlternate"><B>�༭</B></td>
			</tr>
<%
	do while not rs.eof
         Tlink=split(rs(2),"$")
%>
			<tr align=center>
			<td height=25 class=forumrow><font color=red><%=rs("id")%></font></td>
			<td class=forumrow><a href=<%=rs("url")%> target=_blank><%=rs("boardname")%></a></td>
			<td class=forumrow><%if rs("islogo")=1 then%>��<%else%>��<%end if%></td>
			<td class=forumrow><a href="admin_link.asp?action=orders&id=<%=rs("id")%>">����</a>  <a href="admin_link.asp?action=edit&id=<%=rs("id")%>">�༭</a>  <a href="admin_link.asp?action=del&id=<%=rs("id")%>">ɾ��</a></td>
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
	body=body+"<br>"+"���³ɹ������������������"
else
	body=body+"<br>"+"����������������̳��Ϣ��"
end if
end sub

sub savedit()
	set rs= server.createobject ("adodb.recordset")
	sql = "select * from bbslink where id="&request("id")
	rs.Open sql,conn,1,3
	if rs.eof and rs.bof then
	body=body+"<br>"+"����û���ҵ�������̳��"
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
	body=body+"<br>"+"���³ɹ������������������"
end sub

sub del
	dim id
	id = request("id")
	sql="delete from bbslink where id="+id
	conn.Execute(sql)
	body=body+"<br>"+"ɾ���ɹ������������������"
end sub

sub orders()
%><br>
            <table width="95%" border="0" cellspacing="3" cellpadding="0" align=center>
              <tr> 
                <td height="22"><font color="<%=Forum_body(7)%>"><b>������̳��������</b><br>
ע�⣺������Ӧ��̳���������������Ӧ��������ţ�<font color=red>ע�ⲻ�ܺͱ��������̳����ͬ���������</font>��</font>
		</td>
              </tr>
	      <tr>
              <tr bgcolor="<%=Forum_body(3)%>"> 
                <td height="22" align=center><a href="admin_link.asp?action=add"><font color="<%=Forum_body(7)%>">�����µ�������̳</font></a></td>
              </tr>
	      <tr>
		<td><font color="<%=Forum_body(7)%>">
<%
	set rs= server.createobject ("adodb.recordset")
	sql="select * from bbslink where id="&cstr(request("id"))
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "û���ҵ���Ӧ��������̳��"
	else
		response.write "<form action=admin_link.asp?action=updatorders method=post>"
		response.write ""&rs("boardname")&"  <input type=text name=newid size=2 value="&rs("id")&">"
		response.write "<input type=hidden name=id value="&request("id")&">"
		response.write "<input type=submit name=Submit value=�޸�></form>"
	end if
	rs.close
	set rs=nothing
%></font>
		</td>
	      </tr>
              <tr bgcolor="<%=Forum_body(3)%>"> 
                <td height="22" align=center><a href="admin_link.asp?action=add"><font color="<%=Forum_body(7)%>">�����µ�������̳</font></a></td>
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
	response.write "���³ɹ���"
	else
	response.write "����ʧ�ܣ���ָ���˺�����������̳��ͬ����ţ�"
	end if
else
	response.write "����ʧ�ܣ���������ַ����Ϸ������������˺�ԭ����ͬ����ţ�"
end if
end sub
%>