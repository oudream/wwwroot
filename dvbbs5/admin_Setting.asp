<!--#include file =conn.asp-->
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
		call saveconst()
		else
		call consted()
		end if
		conn.close
		set conn=nothing
	end if


sub consted()
dim sel
%>
<form method="POST" action=admin_setting.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height=23 colspan=2 align=left><B>˵��</B>����������ô�����������̳�����е����ã����������úͷ���̳�����������������ü��ײ�ͬ�ķ����Թ�������̳ʹ�á�<b>�����������к��û��鹦���ظ��ģ����û�������Ϊ׼</b></td>
</tr>
<tr> 
<td height=23 colspan=2 align=left>
<table width=85% align=center>
<tr>
<td width=50% ><a name=top></a><a href="#setting1">[�򿪻��߹ر���̳]</a></td>
<td width=50% ><a href="#setting2">[��ˮ���ѡ��]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting3">[������Ϣ]</a></td>
<td width=50% ><a href="#setting4">[��ϵ����]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting5">[����ѡ��]</a></td>
<td width=50% ><a href="#setting6">[���Ļ�ѡ��]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting7">[��̳��ҳѡ��]</a></td>
<td width=50% ><a href="#setting8">[�û���ע��ѡ��]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting9">[ͶƱѡ��]</a></td>
<td width=50% ><a href="#setting10">[������ѡ��(�ű���ʱ�������)]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting11">[�����Ͳ鿴����ѡ��]</a></td>
<td width=50% ><a href="#setting12">[���ߺ��û���Դ]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting13">[�ʼ�ѡ��]</a></td>
<td width=50% ><a href="#setting14">[�ϴ�����]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting15">[�û�ѡ��(ǩ����ͷ�Ρ����е�)]</a></td>
<td width=50% ><a href="#setting16">[�������ѡ��]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting17">[��ˢ�»���]</a></td>
<td width=50% ><a href="#setting18">[��������]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting19">[��������]</a></td>
<td width=50% ></td>
</tr>
</table>
</td>
</tr> 
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2  class="tableHeaderText">��ǰʹ�����ã��ɽ���ǰ���ñ��浽��������ģ���У�ѡ�м��ɣ�
</th></tr>
<tr> 
<td width="100%" colspan=2 class="forumrow">
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
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_setting.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%></font>
</td></tr>
<tr> 
<td width="100%" colspan=2 class="Forumrow"><B>��ǰʹ�ø����õķ���̳</B><BR>���·���̳��ʹ�ñ����ã������޸���̳ʹ�����ã��ɵ���̳��������������<BR>
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
<th height="23" colspan="2" align=left><a name="setting1"><b>�򿪻��߹ر���̳</b></a>[<a href="#top"><FONT color="#FFFFFF">����</FONT></a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>��̳��ǰ״̬</U><BR>ά���ڼ�����ùر���̳ʹ��</td>
<td width="50%" class=Forumrow> 
<input type=radio name="ForumStop" value=0 <%if cint(Forum_Setting(21))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="ForumStop" value=1 <%if cint(Forum_Setting(21))=1 then%>checked<%end if%>>�ر�&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>ά��˵��</U><BR>����̳�ر��������ʾ��֧��html�﷨</td>
<td width="50%" class=Forumrow> 
<textarea name="StopReadme" cols="40" rows="3"><%=StopReadme%></textarea>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��ҳģʽ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="IpFlag" value=0 <%if cint(Forum_Setting(4))=0 then%>checked<%end if%>>�°���ʽ&nbsp;
<input type=radio name="IpFlag" value=1 <%if cint(Forum_Setting(4))=1 then%>checked<%end if%>>�ɰ���ʽ&nbsp;
</td>
</tr>
<tr> 
<th align=left height="23" colspan="2" ><a name="setting2"><b>��ˮ���ѡ��</b></a>[<a href="#top"><FONT color="#FFFFFF">����</font></a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����ˮ����</U><BR>��ѡ�������д��������Ʒ������ʱ��<BR>�԰����͹���Ա��Ч</td>
<td width="50%" class=Forumrow>  
<input type=radio name="relaypost" value=0 <%if cint(Forum_Setting(42))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="relaypost" value=1 <%if cint(Forum_Setting(42))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����ˮʱ����</U><BR>��д����Ŀ��ȷ�������˷���ˮ����</td>
<td width="50%" class=Forumrow>  
<input type="text" name="relaypostTime" size="3" value="<%=Forum_Setting(43)%>">&nbsp;��
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting3"><b>������Ϣ</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳����</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="ForumName" size="35" value="<%=Forum_info(0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳��url</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="ForumURL" size="35" value="<%=Forum_info(1)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��ҳ����</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="CompanyName" size="35" value="<%=Forum_info(2)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��ҳURL</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="HostUrl" size="35" value="<%=Forum_info(3)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳��ҳLogo��ַ</U><BR>��ʾ����̳��ҳ�������̳���û����дlogo��ַ����ʹ�ø�����</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Logo" size="35" value="<%=Forum_info(6)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳ͼƬĿ¼</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Picurl" size="35" value="<%=Forum_info(7)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳����Ŀ¼</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Faceurl" size="35" value="<%=Forum_info(8)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��Ȩ��Ϣ</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Copyright" size="35" value="<%=Copyright%>">
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting4"><b>��ϵ����</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳����ԱEmail</U><BR>���û������ʼ�ʱ����ʾ����ԴEmail��Ϣ</td>
<td width="50%" class=Forumrow>  
<input type="text" name="SystemEmail" size="35" value="<%=Forum_info(5)%>">
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting5"><b>����ѡ��</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����ģʽ</U><BR>�߼�ģʽΪ��ʾ������Ŀ</td>
<td width="50%" class=Forumrow>  
<input type=radio name="TopicFlag" value=0 <%if cint(Forum_Setting(17))=0 then%>checked<%end if%>>��ͨ&nbsp;
<input type=radio name="TopicFlag" value=1 <%if cint(Forum_Setting(17))=1 then%>checked<%end if%>>�߼�&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ���UBB����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="strAllowForumCode" value=0 <%if cint(Forum_Setting(34))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="strAllowForumCode" value=1 <%if cint(Forum_Setting(34))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ���HTML����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="strAllowHTML" value=0 <%if cint(Forum_Setting(35))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="strAllowHTML" value=1 <%if cint(Forum_Setting(35))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ�����ͼ��ǩ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="strIMGInPosts" value=0 <%if cint(Forum_Setting(36))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="strIMGInPosts" value=1 <%if cint(Forum_Setting(36))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ��������ǩ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="strIcons" value=0 <%if cint(Forum_Setting(37))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="strIcons" value=1 <%if cint(Forum_Setting(37))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ���Flash��ǩ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="strflash" value=0 <%if cint(Forum_Setting(38))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="strflash" value=1 <%if cint(Forum_Setting(38))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>�������������ֽ���</U><BR>����������</td>
<td class=Forumrow width="50%"> 
<input type="text" name="AnnounceMaxBytes" size="6" value="<%=Forum_Setting(13)%>">&nbsp;�ֽ�
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�������󷵻�</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="PostRetrun" value=1 <%if Forum_Setting(69)=1 then%>checked<%end if%>>��ҳ&nbsp;
<input type=radio name="PostRetrun" value=2 <%if Forum_Setting(69)=2 then%>checked<%end if%>>��̳&nbsp;
<input type=radio name="PostRetrun" value=3 <%if Forum_Setting(69)=3 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting6"><b>���Ļ�ѡ��</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�¶���Ϣ��������</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="openmsg" value=0 <%if cint(Forum_Setting(10))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="openmsg" value=1 <%if cint(Forum_Setting(10))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting7"><b>��̳��ҳѡ��</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���ٵ�½λ��</U></td>
<td width="50%" class=Forumrow>  
<select name="FastLogin">
<option value="1" <%if cint(Forum_Setting(31))=1 then%>selected<%end if%>>����
<option value="2" <%if cint(Forum_Setting(31))=2 then%>selected<%end if%>>�ײ�
<option value="0" <%if cint(Forum_Setting(31))="0" then%>selected<%end if%>>����ʾ
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ���ʾ�����ջ�Ա</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="BirthFlag" value=0 <%if cint(Forum_Setting(29))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="BirthFlag" value=1 <%if cint(Forum_Setting(29))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting8"><b>�û���ע��ѡ��</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>ͬһIPע����ʱ��</U><BR>�粻�����ƿ���д0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="RegTime" size="3" value="<%=Forum_Setting(22)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>Email֪ͨ����</U><BR>ȷ������վ��֧�ַ���mail������������Ϊϵͳ�������</td>
<td width="50%" class=Forumrow>  
<input type=radio name="EmailReg" value=0 <%if cint(Forum_Setting(23))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="EmailReg" value=1 <%if cint(Forum_Setting(23))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>һ��Emailֻ��ע��һ���ʺ�</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="EmailRegOne" value=0 <%if cint(Forum_Setting(24))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="EmailRegOne" value=1 <%if cint(Forum_Setting(24))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>ע����Ҫ����Ա��֤</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="RegFlag" value=0 <%if cint(Forum_Setting(25))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="RegFlag" value=1 <%if cint(Forum_Setting(25))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����ע����Ϣ�ʼ�</U><BR>��ȷ���������ʼ�����</td>
<td width="50%" class=Forumrow>  
<input type=radio name="SendRegEmail" value=0 <%if cint(Forum_Setting(47))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="SendRegEmail" value=1 <%if cint(Forum_Setting(47))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�������Ż�ӭ��ע���û�</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="smsflag" value=0 <%if cint(Forum_Setting(46))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="smsflag" value=1 <%if cint(Forum_Setting(46))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting9"><b>ͶƱѡ��</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳ͶƱ����Ŀ</U><BR>����д����</td>
<td width="50%" class=Forumrow>  
<input type="text" name="vote_num" size="3" value="<%=Forum_Setting(39)%>">
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting10"><b>������ѡ��</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳����ʱ��</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="GMT" size="35" value="<%=Forum_info(9)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>������ʱ��</U></td>
<td width="50%" class=Forumrow>  
<select name="TimeAdjust">
<%for i=-23 to 23%>
<option value="<%=i%>" <%if cint(i)=cint(Forum_Setting(0)) then%>selected<%end if%>><%=i%>
<%next%>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�ű���ʱʱ��</U><BR>Ĭ��Ϊ300��һ�㲻������</td>
<td width="50%" class=Forumrow>  
<input type="text" name="ScriptTimeOut" size="3" value="<%=Forum_Setting(1)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ���ʾҳ��ִ��ʱ��</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="runtime" value=0 <%if cint(Forum_Setting(30))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="runtime" value=1 <%if cint(Forum_Setting(30))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting11"><b>�����Ͳ鿴����ѡ��</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�����Ƿ�����鿴��Ա����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="ViewUser_g" value=0 <%if cint(Forum_Setting(27))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="ViewUser_g" value=1 <%if cint(Forum_Setting(27))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�û��Ƿ�����鿴��Ա����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="ViewUser_u" value=0 <%if cint(Forum_Setting(28))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="ViewUser_u" value=1 <%if cint(Forum_Setting(28))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ�����δ��½�û�����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Search_G" value=0 <%if cint(Forum_Setting(48))=0 then%>checked<%end if%>>����&nbsp;
<input type=radio name="Search_G" value=1 <%if cint(Forum_Setting(48))=1 then%>checked<%end if%>>������&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting12"><b>���ߺ��û���Դ</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����������ʾ��������</U><BR>Ϊ��ʡ��Դ����ر�</td>
<td width="50%" class=Forumrow>  
<input type=radio name="online_g" value=0 <%if cint(Forum_Setting(15))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="online_g" value=1 <%if cint(Forum_Setting(15))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����������ʾ�û�����</U><BR>Ϊ��ʡ��Դ����ر�</td>
<td width="50%" class=Forumrow>  
<input type=radio name="online_u" value=0 <%if cint(Forum_Setting(14))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="online_u" value=1 <%if cint(Forum_Setting(14))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>ɾ������û�ʱ��</U><BR>������ɾ�����ٷ����ڲ���û�<BR>��λ�����ӣ�����������</td>
<td width="50%" class=Forumrow>  
<input type="text" name="kicktime" size="3" value="<%=Forum_Setting(8)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����ͬʱ������</U><BR>�粻�����ƣ�������Ϊ999999</td>
<td width="50%" class=Forumrow>  
<input type="text" name="online_n" size="6" value="<%=Forum_Setting(26)%>">&nbsp;��
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting13"><b>�ʼ�ѡ��</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�����ʼ����</U><BR>������ķ�������֧�������������ѡ��֧��</td>
<td width="50%" class=Forumrow>  
<select name="EmailFlag">
<option value="0" <%if Forum_Setting(2)=0 then%>selected<%end if%>>��֧�� 
<option value="1" <%if Forum_Setting(2)=1 then%>selected<%end if%>>JMAIL 
<option value="2" <%if Forum_Setting(2)=2 then%>selected<%end if%>>CDONTS 
<option value="3" <%if Forum_Setting(2)=3 then%>selected<%end if%>>ASPEMAIL 
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>SMTP Server��ַ</U><BR>ֻ������̳ʹ�������д��˷����ʼ����ܣ�����д���ݷ���Ч</td>
<td width="50%" class=Forumrow>  
<input type="text" name="SMTPServer" size="35" value="<%=Forum_info(4)%>">
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting14"><b>�ϴ�����</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>ͷ���ϴ�</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="uploadFlag" value=0 <%if Forum_Setting(7)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="uploadFlag" value=1 <%if Forum_Setting(7)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�ϴ��ļ�����</U><BR>ʹ�ö��Ÿ���</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_upload" size="35" value="<%=Forum_upload%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�ϴ��ļ���С</U><BR>����д����</td>
<td width="50%" class=Forumrow>  
<input type="text" name="uploadsize" size="6" value="<%=Forum_Setting(33)%>"> K
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�����ϴ�ͼƬ</U><BR>��������ͼƬ���ϴ�</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Uploadpic" value=0 <%if cint(Forum_Setting(3))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Uploadpic" value=1 <%if cint(Forum_Setting(3))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>����ͼƬ�ϴ�·��</U><BR>�����̳Ŀ¼·��</td>
<td class=Forumrow width="50%"> 
<input type="text" name="UpLoadPath" size="10" value="<%=Forum_Setting(64)%>">&nbsp;�磺uploadFace
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting15"><b>�û�ѡ��</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�û�ǩ���Ƿ���UBB����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="UserubbCode" value=0 <%if Forum_Setting(65)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="UserubbCode" value=1 <%if Forum_Setting(65)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�û�ǩ���Ƿ���HTML����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="UserHtmlCode" value=0 <%if Forum_Setting(66)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="UserHtmlCode" value=1 <%if Forum_Setting(66)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�û��Ƿ�����ͼ��ǩ</U><BR>����ͼƬ��flash����ý���</td>
<td width="50%" class=Forumrow>  
<input type=radio name="UserImgCode" value=0 <%if Forum_Setting(67)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="UserImgCode" value=1 <%if Forum_Setting(67)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>�û������б����</U></td>
<td class=Forumrow width="50%"> 
<input type="text" name="TopUserNum" size="6" value="<%=Forum_Setting(68)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�û�ͷ��</U><BR>�Ƿ������û��Զ���ͷ��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="TitleFlag" value=0 <%if cint(Forum_Setting(6))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="TitleFlag" value=1 <%if cint(Forum_Setting(6))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr><tr> 
<td width="50%" class=Forumrow> <U>�û���ɾ���������Լ���������</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="UserPostAdmin" value=0 <%if Forum_Setting(70)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="UserPostAdmin" value=1 <%if Forum_Setting(70)=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting16"><b>�������ѡ��</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳���</U><BR>���������Ϊ����</td>
<td width="50%" class=Forumrow>  
<input type=radio name="DvbbsSkin" value=0 <%if Forum_Setting(62)=0 then%>checked<%end if%>>Ĭ�Ϸ��&nbsp;
<input type=radio name="DvbbsSkin" value=1 <%if Forum_Setting(62)=1 then%>checked<%end if%>>���������&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>���������Ԥ������</U><BR>ȡ��Ԥ������д0</td>
<td class=Forumrow width="50%"> 
<input type="text" name="SkinFontNum" size="6" value="<%=Forum_Setting(63)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�������ÿҳ��ʾ������</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Maxtitlelist" size="3" value="<%=Forum_Setting(12)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���������ֺ�</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="FontSize" size="3" value="<%=Forum_Setting(59)%>">&nbsp;Pt����15
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���������о�</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="FontHeight" size="3" value="<%=Forum_Setting(60)%>">&nbsp;���о࣬��1.5
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�������ҳ���û�������ʾ</U><BR>һΪ����ʾ����Ʋ���<BR>��Ϊ����ʾ�û�����<BR>��Ϊ���߶�����ʾ����Ϊ���߶���ʾ</td>
<td width="50%" class=Forumrow>  
<input type=radio name="BbsUserInfo" value=1 <%if Forum_Setting(61)=1 then%>checked<%end if%>>ģʽһ&nbsp;
<input type=radio name="BbsUserInfo" value=2 <%if Forum_Setting(61)=2 then%>checked<%end if%>>ģʽ��&nbsp;
<input type=radio name="BbsUserInfo" value=3 <%if Forum_Setting(61)=3 then%>checked<%end if%>>ģʽ��&nbsp;
<input type=radio name="BbsUserInfo" value=4 <%if Forum_Setting(61)=4 then%>checked<%end if%>>ģʽ��&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting17"><b>��ˢ�»���</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��ˢ�»���</U><BR>��ѡ�������д���������ˢ��ʱ��<BR>�԰����͹���Ա��Ч</td>
<td width="50%" class=Forumrow>  
<input type=radio name="ReflashFlag" value=0 <%if cint(Forum_Setting(19))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="ReflashFlag" value=1 <%if cint(Forum_Setting(19))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���ˢ��ʱ����</U><BR>��д����Ŀ��ȷ�������˷�ˢ�»���<BR>���������б����ʾ����ҳ��������</td>
<td width="50%" class=Forumrow>  
<input type="text" name="ReflashTime" size="3" value="<%=Forum_Setting(20)%>">&nbsp;��
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting18"><b>��������</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���������̳</U><BR>�������Ƿ�������������̳</td>
<td width="50%" class=Forumrow>  
<input type=radio name="guestlogin" value=0 <%if cint(Forum_Setting(9))=0 then%>checked<%end if%>>����&nbsp;
<input type=radio name="guestlogin" value=1 <%if cint(Forum_Setting(9))=1 then%>checked<%end if%>>������&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>ÿҳ��ʾ����¼</U><BR>������̳���кͷ�ҳ�йص���Ŀ</td>
<td width="50%" class=Forumrow>  
<input type="text" name="MaxAnnouncePerPage" size="3" value="<%=Forum_Setting(11)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳С�ֱ�</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="SmallPaper" value=0 <%if Forum_Setting(56)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="SmallPaper" value=1 <%if Forum_Setting(56)=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>С�ֱ��Ƿ�Կ��˿���</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="SmallPaper_g" value=0 <%if Forum_Setting(57)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="SmallPaper_g" value=1 <%if Forum_Setting(57)=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����С�ֱ��۳���Ǯ��</U><br>��ѡ����˿ɷ�����ֻ�۳���Ա��Ǯ</td>
<td width="50%" class=Forumrow> 
<input type="text" name="SmallPaper_m" size="3" value="<%=Forum_Setting(58)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����������ɾ������</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="bmflag_1" value=0 <%if cint(Forum_Setting(49))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="bmflag_1" value=1 <%if cint(Forum_Setting(49))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���������޸���ɫ����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="bmflag_2" value=0 <%if cint(Forum_Setting(50))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="bmflag_2" value=1 <%if cint(Forum_Setting(50))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���а��������޸���ɫ����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="bmflag_4" value=0 <%if cint(Forum_Setting(52))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="bmflag_4" value=1 <%if cint(Forum_Setting(52))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>������ɫ����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="viewcolor" value=0 <%if cint(Forum_Setting(55))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="viewcolor" value=1 <%if cint(Forum_Setting(55))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr><tr> 
<td width="50%" class=Forumrow> <U>�Ƿ񹫿������¼�</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="bbsEven" value=0 <%if Forum_Setting(71)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="bbsEven" value=1 <%if Forum_Setting(71)=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ񹫿������¼��еĲ�����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="bbsEvenView" value=0 <%if Forum_Setting(72)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="bbsEvenView" value=1 <%if Forum_Setting(72)=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<th height=23 colspan=2 align=left><a name="setting19"><b>��������</b></a>[<a href="#top"><font color=#FFFFFF>����</font></a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ�����̳����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="GroupFlag" value=0 <%if cint(Forum_Setting(32))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="GroupFlag" value=1 <%if cint(Forum_Setting(32))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>



<tr bgcolor=<%=Forum_body(3)%>> 
<td width="50%" class=Forumrow> &nbsp;</td>
<td width="50%" class=Forumrow>  
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
dim Forum_copyright
if trim(request("skinid"))="" and trim(request("boardid"))="" then
Founderr=true
Errmsg=Errmsg+"<br>"+"<li>��ѡ�񱣴��ģ������"
else
Forum_Setting=request.Form("TimeAdjust") & "," & request.Form("ScriptTimeOut") & "," & request.Form("EmailFlag") & "," & request.Form("Uploadpic") & "," & request.Form("IpFlag") & "," & request.Form("FromFlag") & "," & request.Form("TitleFlag") & "," & request.Form("uploadFlag") & "," & request.Form("kicktime") & "," & request.Form("guestlogin") & "," & request.Form("openmsg") & "," & request.Form("MaxAnnouncePerPage") & "," & request.Form("Maxtitlelist") & "," & request.Form("AnnounceMaxBytes") & "," & request.Form("online_u") & "," & request.Form("online_g") & "," & request.Form("LinkFlag") & "," & request.Form("TopicFlag") & "," & request.Form("VoteFlag") & "," & request.Form("ReflashFlag") & "," & request.Form("ReflashTime") & "," & request.Form("ForumStop") & "," & request.Form("RegTime") & "," & request.Form("EmailReg") & "," & request.Form("EmailRegOne") & "," & request.Form("RegFlag") & "," & request.Form("online_n") & "," & request.Form("ViewUser_g") & "," & request.Form("ViewUser_u") & "," & request.Form("BirthFlag") & "," & request.Form("runtime") & "," & request.Form("FastLogin") & "," & request.Form("GroupFlag") & "," & request.Form("uploadsize") & "," & request.Form("strAllowForumCode") & "," & request.Form("strAllowHTML") & "," & request.Form("strIMGInPosts") & "," & request.Form("strIcons") & "," & request.Form("strflash") & "," & request.Form("vote_num") & "," & request.Form("facenum") & "," & request.Form("imgnum") & "," & request.Form("relaypost") & "," & request.Form("relayposttime") & "," & request.Form("facename") & "," & request.Form("imgname") & "," & request.Form("smsflag") & "," & request.Form("SendRegEmail") & "," & request.Form("Search_G") & "," & request.Form("bmflag_1") & "," & request.Form("bmflag_2") & "," & request.Form("bmflag_3") & "," & request.Form("bmflag_4") & "," & request.Form("bmflag_5") & "," & request.Form("RegFaceNum") & "," & request.Form("viewcolor") & "," & request.Form("SmallPaper") & "," & request.Form("SmallPaper_g") & "," & request.Form("SmallPaper_m")
Forum_Setting=Forum_Setting & "," & request.Form("FontSize") & "," & request.Form("FontHeight") & "," & request.Form("BbsUserInfo") & "," & request.Form("DvbbsSkin") & "," & request.Form("SkinFontNum") & "," & request.Form("UpLoadPath") & "," & request.Form("UserubbCode") & "," & request.Form("UserHtmlCode") & "," & request.Form("UserImgCode") & "," & request.Form("TopUserNum") & "," & request.Form("PostRetrun") & "," & request.Form("UserPostAdmin") & "," & request.Form("bbsEven") & "," & request.Form("bbsEvenView")

Forum_info=Request("ForumName") & "," & Request("ForumURL") & "," & Request("CompanyName") & "," & Request("HostUrl") & "," & Request("SMTPServer") & "," & Request("SystemEmail") & "," & Request("Logo") & "," & Request("Picurl") & "," & Request("Faceurl") & "," & Request("GMT")
Forum_copyright=request("copyright")
if request("skinid")<>"" then
sql="update config set Forum_Setting='"&Forum_Setting&"',StopReadme='"&request.Form("StopReadme")&"',Forum_upload='"&request.Form("Forum_upload")&"',Forum_info='"&Forum_info&"',Forum_copyright='"&Forum_copyright&"' where id in ( "&request("skinid")&" )"
conn.execute(sql)
end if
response.write "���óɹ���"
end if
end sub
%>