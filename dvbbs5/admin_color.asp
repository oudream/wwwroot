<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		if request("action")="save" then
			call saveconst()
		else
			call consted()
		end if
		if founderr then call dvbbs_error()
		conn.close
		set conn=nothing
	end if

sub consted()
dim sel
%>
<form method="POST" action=admin_color.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2" ><B>˵��</B>��<BR>1����ѡ����ѡ���Ϊ��ǰ��ʹ������ģ�壬����ɲ鿴��ģ�����ã�������ģ��ֱ�Ӳ鿴��ģ�岢�޸����á������Խ�����������ñ����ڶ����̳���ģ����<BR>2����Ҳ���Խ������趨����Ϣ���沢Ӧ�õ�����ķ���̳�����У��ɶ�ѡ<BR>3�����������һ���������ñ�İ����ģ������ã�ֻҪ����ð����ģ�����ƣ������ʱ��ѡ��Ҫ���浽�İ������ƻ�ģ�����Ƽ��ɡ�
<BR><B>С֪ʶ</B>��<BR>��CSS�����У�ÿ����÷ֺţ�;���ֿ������õĶ����У�<BR>1��background-color��ʾ������ɫ����background-color: #FFFFFF;
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=3>��ǰʹ����ģ�壨�ɽ����ñ��浽����ģ���У�</th>
</tr>
<tr>
<td colspan=3 class=forumrow>
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
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_color.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%></font>
</td></tr>
<tr> 
<td width="100%" colspan=3 class=forumrow><B>��ǰʹ�ø�ģ��ķ���̳</B><BR>���·���̳��ʹ�ñ�ģ�����ã������޸���̳ʹ��ģ�壬�ɵ���̳��������������<BR>
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
<script>
function color(para_URL){var URL =new String(para_URL)
window.open(URL,'','width=300,height=220,noscrollbars')}
</SCRIPT>
<tr> 
<th height="23" colspan="3" align=left>��̳��������&nbsp;���� <a href="javascript:color('color.asp')"><font color=#FFFFFF>����</font></a> ʹ��������ɫʰȡ�����޸��������ñ���߱�һ����ҳ֪ʶ��</th>
</tr>
<tr> 
<td width="45%" class=forumrow>��̳BODY��ǩ<br>
����������̳���ı�����ɫ���߱���ͼƬ��</td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<input type="text" name="ForumBody" size="50" value="<%=Forum_body(11)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>Body��CSS����<BR>�ɶ�������Ϊ��ҳ������ɫ��������������߿��<BR><FONT COLOR="red">���¶����У�����CSS����ľͲ�һ����������ɫ�ϵ������ˣ��������������е����塢������</FONT></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="iebarcolor" cols="50" rows="3"><%=Forum_body(25)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>��̳�����ܵ�CSS����<BR>�ɶ�������Ϊ����������ɫ����ʽ��<BR></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="ForumCssLink" cols="50" rows="3"><%=Forum_body(6)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>��̳����ܵ�CSS����<BR>����Ϊ��̳�ܵı���壬Ϊһ����ķ�����ã���ҳ���г��˵�������Ϣ�б���Ĳ��֣�β����ͷ���ȣ����ɶ�������Ϊ������������ɫ����ʽ��<BR></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="ForumCssTD" cols="50" rows="3"><%=Forum_body(10)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>�����˵����CSS����(Logo & Banner�Ϸ�)<BR><FONT COLOR="#000099">���ã�Class=TopDarkNav</FONT></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="NavDarkcolor" cols="50" rows="3"><%=Forum_body(24)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>�����˵����CSS����(Logo & Banner�·�)<BR><FONT COLOR="#000099">���ã�Class=TopLighNav</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="Navlighcolor" cols="50" rows="3"><%=Forum_body(23)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>�����˵����CSS����(�����˵�)<BR><FONT COLOR="#000099">���ã�Class=TopLighNav1</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="Navlighcolor1" cols="50" rows="3"><%=Forum_body(26)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>�����˵����CSS����(Logo & banner����)<BR><FONT COLOR="#000099">���ã�Class=TopLighNav2</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="Navlighcolor2" cols="50" rows="3"><%=Forum_body(28)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>���߿���ɫ����<br>
�����涨��һ�Ͷ�CSS������Ʋ����ĵط�����ñ��ֺͱ߿�CSS����һ��ͬ������ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(27)%>"></td>
<td width="50%" class=forumrow> 
<input type=text name="btablebackcolor" value="<%=Forum_body(27)%>" size=35>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>���߿��壩CSS����һ<br>
һ��ҳ�棬�ɶ�������Ϊ����񱳾�������ͼ���߿򡢿�ȵ�<BR><FONT COLOR="#000099">���ã�Class=TableBorder1</font></td>
<td width="5%"  class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="Tablebackcolor" cols="50" rows="3"><%=Forum_body(0)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>���߿��壩CSS�����<br>
����̳�������߿򼰲���ҳ�棬�ɶ�������Ϊ����񱳾�������ͼ���߿򡢿�ȵ�<BR><FONT COLOR="#000099">���ã�Class=TableBorder2</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="aTablebackcolor" cols="50" rows="3"><%=Forum_body(1)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>������CSS����һ�������<br>
һ��Ϊ���ı���״̬����ɫ���߱����ϵĶ��壬�ɶ�������Ϊ����������ͼ�����弰����ɫ��<BR><FONT COLOR="#000099">���ã�th</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="Tabletitlecolor" cols="50" rows="3"><%=Forum_body(2)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>��������������CSS���壨�����<br>
���������Ϊ���������������ر����ڱ����������ӵ���ɫ���������ɫһ�������������ｫ����Ĭ�ϵ�������ɫ<BR>ע�⣺�����#TableTitleLink����ȥ�������Խ�������Ŀ�ֿ�д���ﵽ���ӵĲ�ͬЧ��<BR><FONT COLOR="#000099">���ã�id=TableTitleLink</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="TableTitleLink" cols="50" rows="3"><%=Forum_body(7)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>������CSS�������ǳ������<br>
Ϊ��������״̬����ɫ���߱����ϵĶ��壬����ҳ����̳���������ɶ�������Ϊ����������ͼ�����弰����ɫ��<BR><FONT COLOR="#000099">���ã�Class=TableTitle2</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="aTabletitlecolor" cols="50" rows="3"><%=Forum_body(3)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>�����CSS����һ<BR>�ɶ�������Ϊ����������ͼ�����弰����ɫ��<BR><FONT COLOR="#000099">���ã�Class=TableBody1</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="Tablebodycolor" cols="50" rows="3"><%=Forum_body(4)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>�������ɫ��(1��2��ɫ����ʾ�д���)<BR>�ɶ�������Ϊ����������ͼ�����弰����ɫ��<BR><FONT COLOR="#000099">���ã�Class=TableBody2</font></td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<textarea name="aTablebodycolor" cols="50" rows="3"><%=Forum_body(5)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>�����</td>
<td width="5%" class=forumrow></td>
<td width="50%" class=forumrow> 
<input type="text" name="TableWidth" size="35" value="<%=Forum_body(12)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>��������������ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(8)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="AlertFontColor" size="35" value="<%=Forum_body(8)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>��ʾ���ӵ�ʱ��������ӣ�ת�����ӣ��ظ��ȵ���ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(9)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="ContentTitle" size="35" value="<%=Forum_body(9)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>��ҳ������ɫ<BR>���������</td>
<td width="5%" bgcolor="<%=Forum_body(14)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="BoardLinkColor" size="35" value="<%=Forum_body(14)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>һ���û�����������ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(15)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="user_fc" size="35" value="<%=Forum_body(15)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>һ���û������ϵĹ�����ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(16)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="user_mc" size="35" value="<%=Forum_body(16)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>��������������ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(17)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="bmaster_fc" size="35" value="<%=Forum_body(17)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>���������ϵĹ�����ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(18)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="bmaster_mc" size="35" value="<%=Forum_body(18)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>����Ա����������ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(19)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="master_fc" size="35" value="<%=Forum_body(19)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>���������ϵĹ�����ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(20)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="master_mc" size="35" value="<%=Forum_body(20)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>�������������ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(21)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="vip_fc" size="35" value="<%=Forum_body(21)%>">
</td>
</tr>
<tr> 
<td width="45%" class=forumrow>��������ϵĹ�����ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(22)%>"></td>
<td width="50%" class=forumrow> 
<input type="text" name="vip_mc" size="35" value="<%=Forum_body(22)%>">
</td>
</tr>
<tr> 
<td width="45%" class=Forumrow>&nbsp;</td>
<td width="5%" class=Forumrow></td>
<td width="50%" class=Forumrow> 
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</table>
</form>
<%
end sub

sub saveconst()
dim Forum_body
if trim(request("skinid"))="" and trim(request("boardid"))="" then
Founderr=true
Errmsg=Errmsg+"<br>"+"<li>��ѡ�񱣴��ģ������"
else
Forum_body=request.form("Tablebackcolor") & "|||" & request.form("aTablebackcolor") & "|||" & request.form("Tabletitlecolor") & "|||" & request.form("aTabletitlecolor") & "|||" & request.form("Tablebodycolor") & "|||" & request.form("aTablebodycolor") & "|||" & request.form("ForumCssLink") & "|||" & request.form("TableTitleLink") & "|||" & request.form("AlertFontColor") & "|||" & request.form("ContentTitle") & "|||" & request.form("ForumCssTD") & "|||" & request.form("ForumBody") & "|||" & request.form("TableWidth") & "|||" & request.form("BodyFontColor") & "|||" & request.form("BoardLinkColor") & "|||" & request.form("user_fc") & "|||" & request.form("user_mc") & "|||" & request.form("bmaster_fc") & "|||" & request.form("bmaster_mc") & "|||" & request.form("master_fc") & "|||" & request.form("master_mc") & "|||" & request.form("vip_fc") & "|||" & request.form("vip_mc") & "|||" & request.form("NavLighColor") & "|||" & request.form("NavDarkColor") & "|||" & request.form("IEbarColor") & "|||" & request.form("Navlighcolor1") & "|||" & request.form("btablebackcolor") & "|||" & request.form("Navlighcolor2")
'response.write Forum_body
if request("skinid")<>"" then
sql = "update config set Forum_body='"&Forum_body&"' where id in ( "&request("skinid")&" )"
conn.execute(sql)
end if
if request("boardid")<>"" then
sql = "update board set Forum_body='"&Forum_body&"' where boardid in ( "&request("boardid")&" )"
conn.execute(sql)
end if
response.write "���óɹ���"
end if
end sub
%>