<!--#include file="conn.asp"-->
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
		dim body
		select case request("action")
		case "news"
			call step1()
		case "step2"
			call saveskin()
		case "active"
			call active()
		case "delet"
			call delete()
		case "edit"
			call edit()
		case "savedit"
			call savedit()
		case else
			call skininfo()
		end select
		conn.close
		set conn=nothing
	end if

	sub delete()
	set rs=conn.execute("select skinname from config where active=1 and id="&request("nskinid"))
	if rs.eof and rs.bof then
	conn.execute("delete from config where id="&request("nskinid"))
	response.write "ɾ���ɹ���"
	else
	response.write "������ʾ����ǰʹ�õ���̳ģ�岻��ɾ�����������ø�ģ��Ϊ���ǵ�ǰģ��״̬��"
	end if
	rs.close
	set rs=nothing
	end sub

	sub edit()
	dim trs
	set rs=conn.execute("select skinname,tid from config where id="&request("skinid"))
%>
<form method="POST" action="?action=savedit">
<input type=hidden value="<%=request("skinid")%>" name="skinid">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2" ><B>˵��</B>��<BR>����ϸ��д������Ϣ���޸�ģ�������뷵�ص��������ӡ�
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height="23" colspan="2" >�޸�ģ������</th>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>ģ������</B></td>
<td width="60%" class="forumrow"> 
<input type="text" name="SkinName" size="35" value="<%=rs(0)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow">&nbsp;</td>
<td width="60%" class="forumrow"> 
<input type="submit" name="Submit" value="�ύ�޸�">
</td>
</tr>
</table>
</form>
<%
	rs.close
	set rs=nothing
	end sub

	sub savedit()
	conn.execute("update config set skinname='"&request("skinname")&"' where id="&request("skinid"))
	response.write "�޸ĳɹ���"
	end sub

	sub skininfo()
%>
<form method=post action="admin_skin.asp?action=active">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2" ><b>��̳ģ��ѡ�����</b>��<BR>�ɵ����Ӧģ�����ͳһ�鿴�͹�����������ĳ��ģ��Ϊ��ǰʹ�á�������ص�������Ŀ�ǳ��࣬Ϊ�˲�������������Ե�ǰʹ�õ�ģ���޸ĵ����Ӧ����˵������޸ģ����ģ�����ƽ����޸�ģ������
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr bgcolor="<%=Forum_body(2)%>"> 
<th width="83%" height=22>ģ������</th><th width="17%">��ǰ�Ƿ�ʹ��</th>
</tr>
<%
	set rs=conn.execute("select id,active,skinname from config order by id desc")
	do while not rs.eof
%>
<tr> 
<td height=25 class=forumrow>&nbsp;<a href="admin_skin.asp?action=edit&skinid=<%=rs("id")%>"><%=rs(2)%></a>��<a href="admin_setting.asp?skinid=<%=rs("id")%>"><font color="<%=Forum_body(14)%>"><U>����������Ϣ</U></font></a>  | <a href="admin_color.asp?skinid=<%=rs("id")%>"><font color="<%=Forum_body(14)%>"><U>��������</U></font></a> | <a href="admin_pic.asp?skinid=<%=rs("id")%>"><font color="<%=Forum_body(14)%>"><U>ͼƬ����</U></font></a> | <a href="admin_ads.asp?skinid=<%=rs("id")%>"><font color="<%=Forum_body(14)%>"><U>�������</U></font></a> | <a href="admin_skin.asp?action=delet&nskinid=<%=rs("id")%>" onclick="{if(confirm('ɾ���󲻿ɻָ���ȷ��ɾ����?')){return true;}return false;}"><font color="<%=Forum_body(14)%>"><U>ɾ��</U></font></a>��</td>
<td class=forumrow>&nbsp;<input type=radio name=skinid value="<%=rs(0)%>" <%if rs(1)=1 then%>checked<%end if%>></td>
</tr>
<%
	rs.movenext
	loop
	rs.close
	set rs=nothing
%>
                <tr> 
                  <td width="80%" height=22 align=right class=forumrow><font color="<%=Forum_body(6)%>">&nbsp;��ǰ��̳Ĭ��ʹ��ģ��ֻ��ѡ��һ��&nbsp;</font></td><td width="50%" class=forumrow><input type=submit value="�� ��" name=submit></td>
                </tr>
	       </table>
		   </form>
<%
	end sub

	sub step1()
%>
<form method="POST" action=?action=step2>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2" ><B>˵��</B>��<BR>����ϸ��д������Ϣ��ֻ����д�����в������Ϣ����ģ����ܱ��浽��̳ģ���б��С�
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height="23" colspan="2" >��̳������Ϣ(��ģ��������Ϣ�ڱ���������ɺ��ڹ���ҳ�����޸�)</th>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>ģ������</B></td>
<td width="60%" class="forumrow"> 
<input type="text" name="SkinName" size="35">
</td>
</tr>
<tr> 
<th height="23" colspan="2" >��̳������Ϣ</th>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��̳����</B></td>
<td width="60%" class="forumrow"> 
<input type="text" name="ForumName" size="35" value="<%=Forum_info(0)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��̳��url</B></td>
<td width="60%" class="forumrow"> 
<input type="text" name="ForumURL" size="35" value="<%=Forum_info(1)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��ҳ����</B></td>
<td width="60%" class="forumrow"> 
<input type="text" name="CompanyName" size="35" value="<%=Forum_info(2)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��ҳURL</B></td>
<td width="60%" class="forumrow"> 
<input type="text" name="HostUrl" size="35" value="<%=Forum_info(3)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>SMTP Server��ַ</B><BR>ֻ������̳ʹ�������д��˷����ʼ����ܣ�����д���ݷ���Ч</td>
<td width="60%" class="forumrow"> 
<input type="text" name="SMTPServer" size="35" value="<%=Forum_info(4)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��̳����ԱEmail</B><BR>���û������ʼ�ʱ����ʾ����ԴEmail��Ϣ</td>
<td width="60%" class="forumrow"> 
<input type="text" name="SystemEmail" size="35" value="<%=Forum_info(5)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��̳��ҳLogo��ַ</B><BR>��ʾ����̳��ҳ�������̳���û����дlogo��ַ����ʹ�ø�����</td>
<td width="60%" class="forumrow"> 
<input type="text" name="Logo" size="35" value="<%=Forum_info(6)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��̳ͼƬĿ¼</B></td>
<td width="60%" class="forumrow"> 
<input type="text" name="Picurl" size="35" value="<%=Forum_info(7)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��̳����Ŀ¼</B></td>
<td width="60%" class="forumrow"> 
<input type="text" name="Faceurl" size="35" value="<%=Forum_info(8)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��̳����ʱ��</B></td>
<td width="60%" class="forumrow"> 
<input type="text" name="GMT" size="35" value="<%=Forum_info(9)%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow"><B>��Ȩ��Ϣ</B></td>
<td width="60%" class="forumrow"> 
<input type="text" name="Copyright" size="35" value="<%=Copyright%>">
</td>
</tr>
<tr> 
<td width="40%" class="forumrow">&nbsp;</td>
<td width="60%" class="forumrow"> 
<input type="submit" name="Submit" value="��һ��">
</td>
</tr>
</table>
</form>
<%
	end sub

	sub saveskin()
dim ars
dim Forum_copyright
dim skinname
Forum_info=Request("ForumName") & "," & Request("ForumURL") & "," & Request("CompanyName") & "," & Request("HostUrl") & "," & Request("SMTPServer") & "," & Request("SystemEmail") & "," & Request("Logo") & "," & Request("Picurl") & "," & Request("Faceurl") & "," & Request("GMT")
Forum_copyright=request("copyright")
skinName=request("skinname")

set Ars=conn.execute("select * from config where active=1")

set rs = server.CreateObject ("adodb.recordset")
sql="select * from config"
rs.open sql,conn,1,3
rs.addnew
rs("Forum_body")=ars("Forum_body")
rs("Forum_info")=Forum_info
rs("Forum_copyright")=Forum_copyright
rs("skinname")=skinname
rs("Forum_Setting")=ars("Forum_Setting")
rs("StopReadme")=ars("StopReadme")
rs("Forum_upload")=ars("Forum_upload")
rs("Forum_ubb")=ars("Forum_ubb")
rs("Forum_statepic")=ars("Forum_statepic")
rs("Forum_topicpic")=ars("Forum_topicpic")
rs("Forum_boardpic")=ars("Forum_boardpic")
rs("Forum_pic")=ars("Forum_pic")
rs("Forum_ads")=ars("Forum_ads")
rs("Forum_user")=ars("Forum_user")
rs("badwords")=badwords
rs("splitwords")=splitwords
rs("birthuser")=ars("birthuser")
rs("Maxonline")=ars("Maxonline")
rs("MaxonlineDate")=ars("MaxonlineDate")
rs("LastPost")="$$"&Now()&"$$$$"
rs("active")=0
rs("allposttable")=ars("allposttable")
rs("allposttablename")=ars("allposttablename")
rs("NowUseBBS")=ars("NowUseBBS")
rs("Cookiepath")=ars("Cookiepath")
rs("tid")=request("tskinid")
rs.update
rs.close
set rs=nothing
response.write "ģ����ӳɹ���"
	end sub

	sub active()
	set rs=conn.execute("select * from config where active=1")
	conn.execute("update config set active=1,Maxonline="&rs("Maxonline")&",MaxonlineDate='"&rs("MaxonlineDate")&"',TopicNum="&rs("TopicNum")&",BbsNum="&rs("BbsNum")&",TodayNum="&rs("TodayNum")&",UserNum="&rs("UserNum")&",lastUser='"&rs("lastUser")&"',yesterdaynum="&rs("yesterdaynum")&",maxpostnum="&rs("maxpostnum")&",Cookiepath='"&rs("Cookiepath")&"' where id="&request("skinid"))
	conn.execute("update config set active=0 where id="&rs("id"))
	response.write "���³ɹ���<a href=admin_skin.asp>����</a>��"
	rs.close
	set rs=nothing
	end sub
%>