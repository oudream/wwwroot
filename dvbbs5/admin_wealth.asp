<!--#include file=conn.asp-->
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
		if request("action")="save" then
		call savegrade()
		else
		call grade()
		end if
		if founderr then call dvbbs_error()
		conn.close
		set conn=nothing
	end if

sub grade()
dim sel
%>
<form method="POST" action=admin_wealth.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2" ><B>˵��</B>��<BR>1����ѡ����ѡ���Ϊ��ǰ��ʹ������ģ�壬����ɲ鿴��ģ�����ã�������ģ��ֱ�Ӳ鿴��ģ�岢�޸����á������Խ�����������ñ����ڶ����̳���ģ����<BR>2����Ҳ���Խ������趨����Ϣ���沢Ӧ�õ�����ķ���̳�����У��ɶ�ѡ<BR>3�����������һ���������ñ�İ����ģ������ã�ֻҪ����ð����ģ�����ƣ������ʱ��ѡ��Ҫ���浽�İ������ƻ�ģ�����Ƽ��ɡ�</font></td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2><B>��ǰʹ����ģ��</B>���ɽ����ñ��浽����ģ���У�</th>
</tr>
<tr>
<td colspan=2 class=forumrow>
<%
set rs = server.CreateObject ("adodb.recordset")
sql="select * from config"
rs.open sql,conn,1,1
do while not rs.eof
if request("skinid")="" then
	if request("boardid")="" then
	if rs("active")=1 then
	sel="checked"
	else
	sel=""
	end if
	else
	sel=""
	end if
else
	if rs("id")=cint(request("skinid")) then
	sel="checked"
	skinid=rs("id")
	else
	sel=""
	end if
end if
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_wealth.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%>
</td></tr>
<tr> 
<td width="100%" colspan=2 class=Forumrow><B>��ǰʹ�ø�ģ��ķ���̳</B><BR>���·���̳��ʹ�ñ�ģ�����ã������޸���̳ʹ��ģ�壬�ɵ���̳��������������<BR>
<select size=1>
<%
set rs = server.CreateObject ("adodb.recordset")
sql="select * from board where sid="&skinid
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "<option>û����̳ʹ�ø�ģ��</option>"
else
do while not rs.eof
response.write "<option>" & rs("boardtype") & "</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%></select>
</td></tr>
<tr> 
<th height="23" colspan="2">�û���Ǯ�趨</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>ע���Ǯ��</td>
<td width="60%" class=Forumrow> 
<input type="text" name="wealthReg" size="35" value="<%=Forum_user(0)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��½���ӽ�Ǯ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="wealthLogin" size="35" value="<%=Forum_user(4)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������ӽ�Ǯ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="wealthAnnounce" size="35" value="<%=Forum_user(1)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������ӽ�Ǯ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="wealthReannounce" size="35" value="<%=Forum_user(2)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������ӽ�Ǯ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="BestWealth" size="35" value="<%=Forum_user(15)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>ɾ�����ٽ�Ǯ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="wealthDel" size="35" value="<%=Forum_user(3)%>">
</td>
</tr>
<tr> 
<th height="23" colspan="2" >�û������趨</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>ע�ᾭ��ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="epReg" size="35" value="<%=Forum_user(5)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��½���Ӿ���ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="epLogin" size="35" value="<%=Forum_user(9)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������Ӿ���ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="epAnnounce" size="35" value="<%=Forum_user(6)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������Ӿ���ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="epReannounce" size="35" value="<%=Forum_user(7)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������Ӿ���ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="bestuserep" size="35" value="<%=Forum_user(17)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>ɾ�����پ���ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="epDel" size="35" value="<%=Forum_user(8)%>">
</td>
</tr>
<tr> 
<th height="23" colspan="2" >�û������趨</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>ע������ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="cpReg" size="35" value="<%=Forum_user(10)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��½��������ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="cpLogin" size="35" value="<%=Forum_user(14)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>������������ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="cpAnnounce" size="35" value="<%=Forum_user(11)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>������������ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="cpReannounce" size="35" value="<%=Forum_user(12)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>������������ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="bestusercp" size="35" value="<%=Forum_user(16)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>ɾ����������ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="cpDel" size="35" value="<%=Forum_user(13)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>&nbsp;</td>
<td width="60%" class=Forumrow> 
<div align="center"> 
<input type="submit" name="Submit" value="�� ��">
</div>
</td>
</tr>
</table>
</form>
<%
end sub

sub savegrade()
if trim(request("skinid"))="" and trim(request("boardid"))="" then
Founderr=true
Errmsg=Errmsg+"<br>"+"<li>��ѡ�񱣴��ģ������"
else
Forum_user=request.form("wealthReg") & "," & request.form("wealthAnnounce") & "," & request.form("wealthReannounce") & "," & request.form("wealthDel") & "," & request.form("wealthLogin") & "," & request.form("epReg") & "," & request.form("epAnnounce") & "," & request.form("epReannounce") & "," & request.form("epDel") & "," & request.form("epLogin") & "," & request.form("cpReg") & "," & request.form("cpAnnounce") & "," & request.form("cpReannounce") & "," & request.form("cpDel") & "," & request.form("cpLogin") & "," & request.form("BestWealth") & "," & request.form("BestuserCP") & "," & request.form("BestuserEP")
'response.write Forum_user
if request("skinid")<>"" then
sql = "update config set Forum_user='"&Forum_user&"' where id in ( "&request("skinid")&" )"
conn.execute(sql)
end if
response.write "���óɹ���"
end if
end sub
%>