<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
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
<form method="POST" action=admin_ads.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2"><B>˵��</B>��<BR>1����ѡ����ѡ���Ϊ��ǰ��ʹ������ģ�壬����ɲ鿴��ģ�����ã�������ģ��ֱ�Ӳ鿴��ģ�岢�޸����á������Խ�����������ñ����ڶ����̳���ģ����<BR>2����Ҳ���Խ������趨����Ϣ���沢Ӧ�õ�����ķ���̳�����У��ɶ�ѡ<BR>3�����������һ���������ñ�İ����ģ������ã�ֻҪ����ð����ģ�����ƣ������ʱ��ѡ��Ҫ���浽�İ������ƻ�ģ�����Ƽ��ɡ�</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2 class="tableHeaderText"><B>��ǰʹ����ģ��</B>���ɽ����ñ��浽����ģ���У�
</th>
</tr>
<tr>
<td colspan=2 class="forumrow">
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
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_ads.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%></font>
</td></tr>
<tr> 
<td width="100%" colspan=2  class="forumrow"><B>��ǰʹ�ø�ģ��ķ���̳</B><BR>���·���̳��ʹ�ñ�ģ�����ã������޸���̳ʹ��ģ�壬�ɵ���̳��������������<BR>
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
<th height="23" colspan="2" class="tableHeaderText"><b>��̳�������</b>����Ϊ���÷���̳�����Ƿ���̳��ҳ��棬����ҳ��Ϊ������ʾҳ�棩</th>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��ҳ����������</B></font></td>
<td width="60%" class="forumrow"> 
<textarea name="index_ad_t" cols="50" rows="3"><%=Forum_ads(0)%></textarea>
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��ҳβ��������</B></font></td>
<td width="60%" class="forumrow"> 
<textarea name="index_ad_f" cols="50" rows="3"><%=Forum_ads(1)%></textarea>
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>������ҳ�������</B></font></td>
<td width="60%" class="forumrow"> 
<input type=radio name="index_moveFlag" value=0 <%if Forum_ads(2)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="index_moveFlag" value=1 <%if Forum_ads(2)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��̳��ҳ�������ͼƬ��ַ</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="MovePic" size="35" value="<%=Forum_ads(3)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��̳��ҳ����������ӵ�ַ</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="MoveUrl" size="35" value="<%=Forum_ads(4)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��̳��ҳ�������ͼƬ���</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="move_w" size="3" value="<%=Forum_ads(5)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��̳��ҳ�������ͼƬ�߶�</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="move_h" size="3" value="<%=Forum_ads(6)%>">&nbsp;����
</td>
</tr>
<input type=hidden name="Board_moveFlag" value=0>
<tr> 
<td width="40%" class="forumrow"><B>������ҳ���¹̶����</B></font></td>
<td width="60%" class="forumrow"> 
<input type=radio name="index_fixupFlag" value=0 <%if Forum_ads(13)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="index_fixupFlag" value=1 <%if Forum_ads(13)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��̳��ҳ���¹̶����ͼƬ��ַ</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="fixupPic" size="35" value="<%=Forum_ads(8)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��̳��ҳ���¹̶�������ӵ�ַ</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="fixupUrl" size="35" value="<%=Forum_ads(9)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��̳��ҳ���¹̶����ͼƬ���</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="fixup_w" size="3" value="<%=Forum_ads(10)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��̳��ҳ���¹̶����ͼƬ�߶�</B></font></td>
<td width="60%" class="forumrow"> 
<input type="text" name="fixup_h" size="3" value="<%=Forum_ads(11)%>">&nbsp;����
</td>
</tr>
<input type=hidden name="Board_fixupFlag" value=0>
<tr> 
<td width="40%" class="forumrow">&nbsp;</td>
<td width="60%" class="forumrow"> 
<div align="center"> 
<input type="submit" name="Submit" value="�� ��">
</div>
</td>
</tr>
</table>
</form>
<%
end sub

sub saveconst()
if trim(request("skinid"))="" and trim(request("boardid"))="" then
Founderr=true
Errmsg=Errmsg+"<br>"+"<li>��ѡ�񱣴��ģ������"
else
Forum_ads=request("index_ad_t") & "$" & request("index_ad_f") & "$" & request("index_moveFlag") & "$" & request("MovePic") & "$" & request("MoveUrl") & "$" & request("move_w") & "$" & request("move_h") & "$" & request("Board_moveFlag") & "$" & request("fixupPic") & "$" & request("FixupUrl") & "$" & request("Fixup_w") & "$" & request("Fixup_h") & "$" & request("Board_fixupFlag") & "$" & request("index_fixupFlag")
'response.write Forum_ads
if request("skinid")<>"" then
sql = "update config set Forum_ads='"&CheckStr(Forum_ads)&"' where id in ( "&request("skinid")&" )"
conn.execute(sql)
end if
response.write "���óɹ���"
end if
end sub
%>