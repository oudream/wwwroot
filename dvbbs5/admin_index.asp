<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="md5.asp" -->
<%
dim username
dim password
dim ip
select case request("action")
case "admin_left"
	call admin_left()
case "admin_login"
	call admin_login()
case "admin_main"
	call admin_main()
case "admin_head"
	call admin_head()
case else
	call main()
end select
sub main()
%>
<html>
<head>
<title><%=Forum_info(0)%>--�������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<frameset cols="180,*" frameborder="NO" border="0" framespacing="0" rows="*"> 
  <frame name="leftFrame" scrolling="AUTO" noresize src="admin_index.asp?action=admin_left" marginwidth="0" marginheight="0">

<frameset rows="20,*"  framespacing="0" border="0" frameborder="0" frameborder="no" border="0">
<frame src="admin_index.asp?action=admin_head" name="head" scrolling="NO" NORESIZE frameborder="0" marginwidth="10" marginheight="0" border="no">
<%if not master or session("flag")="" then%>
  <frame name="main" src="admin_index.asp?action=admin_login" scrolling="AUTO" NORESIZE frameborder="0" marginwidth="10" marginheight="10" border="no">
<%else%>
  <frame name="main" src="admin_index.asp?action=admin_main" scrolling="AUTO" NORESIZE frameborder="0" marginwidth="10" marginheight="10" border="no">
<%end if%>
</frameset>
</frameset>
<noframes>

</body></noframes>
</html>
<%
end sub

sub admin_left()
%>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<%
REM ������Ŀ����
dim menu(6,10)
menu(0,0)="��������"
menu(0,1)="<a href=admin_setting.asp target=main>����������Ϣ</a>"
menu(0,2)="<a href=admin_ads.asp target=main>��̳�������</a>"
menu(0,3)="<a href=bbseven.asp target=main>������־</a>"
menu(0,4)="<a href=announcements.asp?boardid=0&action=AddAnn target=_top>������̳����</a>"

menu(1,0)="��̳����"
menu(1,1)="<a href=admin_board.asp?action=addclass target=main>�����̳����</a>"
menu(1,2)="<a href=admin_board.asp?action=add target=main>��̳�������</a> | <a href=admin_board.asp target=main>����</a>"
menu(1,3)="<a href=admin_board.asp?action=permission target=main>��̳Ȩ�޹���</a>"
menu(1,4)="<a href=admin_boardunite.asp target=main>�ϲ���̳����</a>"
menu(1,5)="<a href=admin_update.asp target=main>������̳����</a>"
menu(1,6)="<a href=admin_link.asp?action=add target=main>������̳���</a> | <a href=admin_link.asp target=main>����</a>"

menu(2,0)="�������������"
menu(2,1)="<a href=admin_alldel.asp target=main>����ɾ��</a>"
menu(2,2)="<a href=admin_alldel.asp?action=moveinfo target=main>�����ƶ�</a>"
menu(2,3)="<a href=# title=Ԥ����δ�ṩ������>�������</a>"
menu(2,4)="<a href=recycle.asp target=main>����վ����</a>"
menu(2,5)="<a href=admin_postdata.asp?action=Nowused target=main>��ǰ�������ݱ����</a>"
menu(2,6)="<a href=admin_postdata.asp target=main>���ݱ������ת��</a>"

menu(3,0)="�û�����"
menu(3,1)="<a href=admin_user.asp target=main>�û���Ϣ����</a>"
menu(3,2)="<a href=admin_grade.asp?action=add target=main>��̳�ȼ����</a> | <a href=admin_grade.asp target=main>�޸�</a>"
menu(3,3)="<a href=admin_wealth.asp target=main>�û���������</a>"
menu(3,4)="<a href=admin_message.asp target=main>��̳���Ź���</a>"
menu(3,5)="<a href=admin_group.asp?action=addgroup target=main>�û������</a> | <a href=admin_group.asp target=main>�޸�</a>"
menu(3,6)="<a href=admin_admin.asp?action=add target=main>����Ա���</a> | <a href=admin_admin.asp target=main>����</a>"
menu(3,7)="<a href=admin_mailist.asp target=main>�ʼ��б�</a>"
menu(3,8)="<a href=admin_update.asp?action=updateuser target=main>�����û�����</a>"

menu(4,0)="�������"
menu(4,1)="<a href=admin_color.asp target=main>��̳���CSS����</a>"
menu(4,2)="<a href=admin_pic.asp target=main>����ͼƬ����</a>"
menu(4,3)="<a href=admin_skin.asp?action=news target=main>��̳����ģ�����</a> | <a href=admin_skin.asp target=main>����</a>"
menu(4,4)="<a href=admin_loadskin.asp target=main>CSS�����</a>"

menu(5,0)="�滻/���ƴ���"
menu(5,1)="<a href=admin_badword.asp?reaction=badword target=main>���ӹ����ַ�</a>"
menu(5,2)="<a href=admin_badword.asp?reaction=splitreg target=main>ע������ַ�</a>"
menu(5,3)="<a href=admin_lockip.asp?action=add target=main>IP�����޶����</a> | <a href=admin_lockip.asp target=main>�޸�</a>"
menu(5,4)="<a href=admin_address.asp?action=add target=main>��̳IP�����</a> | <a href=admin_address.asp target=main>�޸�</a>"
menu(5,5)="<a href=admin_address.asp?action=upip target=main>����IP��</a>"

menu(6,0)="���ݴ���(Access)"
menu(6,1)="<a href=admin_CompressData.asp target=main>ѹ�����ݿ�</a>"
menu(6,2)="<a href=admin_BackupData.asp target=main>�������ݿ�</a>"
menu(6,3)="<a href=admin_RestoreData.asp target=main>�ָ����ݿ�</a>"
menu(6,4)="<a href=admin_SpaceSize.asp target=main>ϵͳ�ռ�ռ��</a>"
%>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table border="0" cellspacing="0" width="100%"  cellpadding="0">
<tr> 
<th height="25" align=center><a href=admin_index.asp target=_top><font color="#FFFFFF"><b>��̳��������</b></font></a></th>
</tr>
<%
	dim j
	dim tmpmenu
	dim menuname
	dim menurl
for i=0 to ubound(menu,1)
%>
<tr> 
<td height="23" class="forumHeaderBackgroundAlternate">&nbsp;<b><%=menu(i,0)%><b></td>
</tr>
	<%
	for j=1 to ubound(menu,2)
	if isempty(menu(i,j)) then exit for
	%>
<tr> 
<td height="20" class="forumrow">&nbsp;&nbsp;<%=menu(i,j)%></td>
</tr>
	<%next%>
<tr> 
<td height="5" class="forumrow"></td>
</tr>
<%next%>
<tr> 
<td height="20" class="forumHeaderBackgroundAlternate">&nbsp; <b><a href=admin_logout.asp target=_top>�˳�����</a></b></td>
</tr>
<tr> 
<td height="20" class="forumrow">&nbsp;<a href=http://www.aspsky.net><font face=Verdana, Arial, Helvetica, 

sans-serif size=1><b>AspSky<font color=#CC0000>.Net</font></b></font></a></td>
</tr>
</table>
<%
end sub

sub admin_login()
	stats="��̳�����½"
	if not founduser then
		Errmsg="<li>������ϵͳ����Ա��"
		founderr=true
	end if
	if founderr then
		call nav()
		call head_var(1)
		call dvbbs_error()
		call footer()
	else
		if request("reaction")="chklogin" then
			call chklogin()
			if founderr then
				call nav()
				call head_var(1)
				call Dvbbs_error()
				call footer()
			end if
		else
			dim num1
			dim rndnum
			Randomize
			Do While Len(rndnum)<4
			num1=CStr(Chr((57-48)*rnd+48))
			rndnum=rndnum&num1
			loop
			session("verifycode")=rndnum
			call admin_login_main()
		end if
	end if
end sub

sub admin_login_main()
%>
<html>
<head>
<meta NAME=GENERATOR Content="Microsoft FrontPage 4.0" CHARSET=GB2312>
<meta name=keywords content='�����ȷ�,������̳,dvbbs'>
<title><%=Forum_info(0)%>--<%=stats%></title>
<link rel=stylesheet href="Forum_admin.css" type=text/css>
</head>
<body leftmargin=0 bottommargin=0 rightmargin=0 topmargin=0 marginheight=0 marginwidth=0>
<title><%=Forum_info(0)%>--<%=stats%></title>
<p>&nbsp;</p>
<form action="admin_index.asp?action=admin_login&reaction=chklogin" method=post> 
<table width=95% border=0 cellspacing=1 cellpadding=3  align=center class=tableBorder>
    <tr>
    <th bgcolor=<%=Forum_body(2)%> valign=middle colspan=2>�����������û����������½������������ܰ��������½��</th></tr>
    <tr>
    <td valign=middle class=Forumrow>�����������û���</td>
    <td valign=middle class=Forumrow><INPUT name=username type=text> &nbsp; <a href="reg.asp">û��ע�᣿</a></td></tr>
    <tr>
    <td valign=middle class=Forumrow>��������������</font></td>
    <td valign=middle class=Forumrow><INPUT name=password type=password> &nbsp; <a href="lostpass.asp">�������룿</a></td></tr>
    <tr>
    <td valign=middle class=Forumrow>���������ĸ�����</td>
    <td valign=middle class=Forumrow><INPUT name=verifycode type=text> &nbsp; ���ڸ����������  <b><span style="background-color: #FFFFFF"><font color=#000000><%=session("verifycode")%></font></span></b></td></tr>
    <tr>
    <td valign=middle colspan=2 align=center class=forumRow><input type=submit name=submit value="�� ½"></td></tr></table></form>
</body>
</html>
<%
end sub

sub chklogin()
	username=trim(replace(request("username"),"'",""))
	password=md5(trim(replace(request("password"),"'","")))
	if request("verifycode")="" then
		errmsg=errmsg+"<br>"+"<li>�뷵������ȷ���롣"
		founderr=true
	elseif session("verifycode")="" then
		errmsg=errmsg+"<br>"+"<li>�벻Ҫ�ظ��ύ���������µ�½�뷵�ص�½ҳ�档"
		founderr=true
	elseif session("verifycode")<>trim(request("verifycode")) then
		errmsg=errmsg+"<br>"+"<li>�������ȷ�����ϵͳ�����Ĳ�һ�£����������롣"
		founderr=true
	end if
	session("verifycode")=""
	if username="" or password="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�����������û��������롣"
	end if
	if founderr then exit sub
	ip=Request.ServerVariables("REMOTE_ADDR")
	set rs=conn.execute("select * from admin where username='"&username&"' and adduser='"&membername&"'")
	if rs.eof and rs.bof then
		rs.close
		set rs=nothing
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��������û��������벻��ȷ����������ϵͳ����Ա��<br><li>��<a href=admin_login.asp>��������</a>�������롣"
		exit sub
	else
		if trim(rs("password"))<>password then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��������û��������벻��ȷ����������ϵͳ����Ա��<br><li>��<a href=admin_login.asp>��������</a>�������롣"
		exit sub
		else
		session("flag")=username
		session.timeout=45
		conn.execute("update admin set LastLogin=Now(),LastLoginIP='"&ip&"' where username='"&username&"'")
		rs.close
		set rs=nothing
		response.write "<script>location.href='admin_index.asp?action=admin_main'</script>"
		end if
	end if
end sub

sub admin_main()
%>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣"
	call dvbbs_error()
else
	Dim theInstalledObjects(17)
    theInstalledObjects(0) = "MSWC.AdRotator"
    theInstalledObjects(1) = "MSWC.BrowserType"
    theInstalledObjects(2) = "MSWC.NextLink"
    theInstalledObjects(3) = "MSWC.Tools"
    theInstalledObjects(4) = "MSWC.Status"
    theInstalledObjects(5) = "MSWC.Counters"
    theInstalledObjects(6) = "IISSample.ContentRotator"
    theInstalledObjects(7) = "IISSample.PageCounter"
    theInstalledObjects(8) = "MSWC.PermissionChecker"
    theInstalledObjects(9) = "Scripting.FileSystemObject"
    theInstalledObjects(10) = "adodb.connection"
    
    theInstalledObjects(11) = "SoftArtisans.FileUp"
    theInstalledObjects(12) = "SoftArtisans.FileManager"
    theInstalledObjects(13) = "JMail.SMTPMail"
    theInstalledObjects(14) = "CDONTS.NewMail"
    theInstalledObjects(15) = "Persits.MailSender"
    theInstalledObjects(16) = "LyfUpload.UploadFile"
    theInstalledObjects(17) = "Persits.Upload.1"
%>
<table cellpadding=3 cellspacing=1 border=0 width=95% align=center>
<tr>
<td width="100%" valign=top>
<p><b>��ӭ���� ������̳ ����������</b>&nbsp;<%=Version%><br><BR>
�ٷ���վ��<a href="http://www.aspsky.net/" target="_blank">�����ȷ�</a><br>
�ٷ���̳��<a href="http://www.dvbbs.net/" target="_blank">������̳</a>
</p>
����������Կ��������е���̳���á����ڴ�ҳ�����ѡ����Ҫ���й�������ӡ�<br>
������ʾ���û�Ȩ�޵�����˳��Ϊ �û���Ĭ�����ã�����̳�û���Ȩ�����ã���(��)��̳�û�Ȩ������
</td>
</tr>
</table>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>��̳��Ϣͳ��</th><tr>
<tr>
<td width="50%"  class="forumRow" height=23>���������ͣ�<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
<td width="50%" class="forumRowHighlight">�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
</tr>
<tr>
<td width="50%" class="forumRow" height=23>վ������·����<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
<td width="50%" class="forumRowHighlight">���ݿ��ַ��</td>
</tr>
<tr>
<td width="50%" class="forumRow" height=23>FSO�ı���д��<%If Not IsObjInstalled(theInstalledObjects(9)) Then%><font color="<%=Forum_body(8)%>"><b>��</b></font><%else%><b>��</b><%end if%></td>
<td width="50%" class="forumRowHighlight">���ݿ�ʹ�ã�<%If Not IsObjInstalled(theInstalledObjects(10)) Then%><font color="<%=Forum_body(8)%>"><b>��</b></font><%else%><b>��</b><%end if%></td>
</tr>
<tr>
<td width="50%" class="forumRow" height=23>Jmail���֧�֣�<%If Not IsObjInstalled(theInstalledObjects(13)) Then%><font color="<%=Forum_body(8)%>"><b>��</b></font><%else%><b>��</b><%end if%></td>
<td width="50%" class="forumRowHighlight">CDONTS���֧�֣�<%If Not IsObjInstalled(theInstalledObjects(14)) Then%><font color="<%=Forum_body(8)%>"><b>��</b></font><%else%><b>��</b><%end if%></td>
</tr>
</table>

<p></p>

<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>��̳�����ݷ�ʽ</th><tr>
<FORM METHOD=POST ACTION="admin_user.asp?action=userSearch&userSearch=9"><tr>
<td width="20%"  class="forumRow" height=23>���ٲ����û�</td>
<td width="80%" class="forumRowHighlight">
<input type="text" name="username" size="30"><input type="submit" value="���̲���">
<input type="hidden" name="userclass" value="0">
<input type="hidden" name="searchMax" value=100>
</td></FORM>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>���õ�����</td>
<td width="80%" class="forumRowHighlight"><select onchange="jumpto(this.options[this.selectedIndex].value)">
		<option>>> ���õ����� <<</option>
		<option value="http://www.dvbbs.net/">��������֧����̳</option>
		<option value="http://www.aspsky.net/article/">�����ȷ�������</option>
		<option value="http://www.dvbbs.net/vbscript/vbstoc.htm">Vbscript���߰����ĵ�</option>
</select></td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>��ݹ�������</td>
<td width="80%" class="forumRowHighlight"><a href=admin_board.asp?action=addclass>�����̳���</a> | <a href=admin_board.asp?action=add>�����̳����</a></td>
</tr>
<tr><form action="admin_update.asp?action=updat" method=post>
<td width="20%" class="forumRow" height=23>���ٸ�������</td>
<td width="80%" class="forumRowHighlight">
<input type="submit" name="Submit" value="������̳����">&nbsp;
<input type="submit" name="Submit" value="������̳������">
</td></form>
</tr>
</table>

<p></p>

<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>������̳���� & ����</th><tr>
<tr>
<td width="20%"  class="forumRow" height=23>��Ʒ����</td>
<td width="80%" class="forumRowHighlight">
<a href="http://www.dvbbs.net/dispuser.asp?name=ɳ̲С��">ɳ̲С��</a>����ϵ��13088878365��QQ:421634��eway@aspsky.net��
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>��ҵ��չ</td>
<td width="80%" class="forumRowHighlight"><a href="http://www.dvbbs.net/dispuser.asp?name=quest">quest</a>����ϵ��13005069033��QQ:398196��quest@aspsky.net��</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>������Ա</td>
<td width="80%" class="forumRowHighlight"><a href="http://www.dvbbs.net/dispuser.asp?name=ɳ̲С��">ɳ̲С��</a>��<a href="http://www.dvbbs.net/dispuser.asp?name=ľ��">ľ��</a>��<a href="http://www.dvbbs.net/dispuser.asp?name=fssunwin">fssunwin</a></td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>�������</td>
<td width="80%" class="forumRowHighlight">
<font color=gray>����</font>
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>�������</td>
<td width="80%" class="forumRowHighlight">
<font color=gray>����</font>
</td>
</tr>
</table>


<table align=center><BR>
Copyright (c) 2001-2002 <a href=http://www.aspsky.net><font face=Verdana, Arial, Helvetica, sans-serif size=1><b>AspSky<font color=#CC0000>.Net</font></b></font></a>. All Rights Reserved .<BR><BR><BR>
</td>
            </tr>
        </table>
<%
end if
end sub

Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function

sub admin_head()
%>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<table width="100%" align=center cellpadding=0 cellspacing=0 border=0>
<tr>
<td height="23" class="forumHeaderBackgroundAlternate" valign=middle>��>>��ӭ<%=membername%>����������</td>
<td class="forumHeaderBackgroundAlternate" valign=middle><a href="http://www.dvbbs.net/" target=_blank>������Ա����</a></td>
<td class="forumHeaderBackgroundAlternate" align=right valign=middle><a href="index.asp" target=_top>������̳��ҳ</a>����</td>
</tr>
</table>
<%
end sub
%>