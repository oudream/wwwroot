<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/email.asp"-->
<!--#include file="md5.asp"-->
<%
dim topic,mailbody,sendmsg,SendMail,useremail
dim username,password,repassword
dim answer
stats="��������"
call nav()
call head_var(1)
if founderr then
	call dvbbs_error()
else
	if request("action")="step1" then
		call step1()
	elseif request("action")="step2" then
		call step2()
	elseif request("action")="step3" then
		call step3()
	else
		call main()
	end if
	if founderr then call dvbbs_error()
end if
call footer()

sub step1()
if request("username")="" then
	Founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�����������û�����"
	exit sub
else
	username=replace(request("username"),"'","")
end if
if Forum_Setting(2)>0 then
	set rs=conn.execute("Select Quesion,Answer,Username,Usergroupid from [user] where username='"&username&"'")
else
	set rs=conn.execute("Select Quesion,Answer,Username,Usergroupid from [user] where username='"&username&"' and UserGroupID>3")
end if
if rs.eof and rs.bof then
	Founderr=true
	errmsg=Errmsg+"<br>"+"<li>��������û����������ڣ����������롣<li>�������ڸ�ϵͳ��֧���ʼ����ͣ������ĵȼ�Ϊ�������ϼ���ֻ��ͨ����ϵ����Ա������롣"
elseif rs(3) < 4 then
	Founderr=true
	errmsg=Errmsg+"<br>"+"<li>���ǹ���Ա�����ǰ�����ֻ��ͨ����ϵ����Ա������롣"
else
	if rs(0)="" or isnull(rs(0)) then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>���û�û����д�������⼰�𰸣�ֻ����д���û����ܼ�����"
	else
%>
<form action="lostpass.asp?action=step2" method="post"> 
<table cellpadding=1 cellspacing=1 align=center class=tableborder1>
    <tr>
    <th valign=middle colspan=2 align=center height=25>ȡ�����루�ڶ������ش����⣩</th></tr>
    <tr>
    <td valign=middle class=tablebody1><b>�ʣ��⣺</b></td>
    <td class=tablebody1 valign=middle><%=rs(0)%></td>
    </tr>
    <tr>
    <td class=tablebody1 valign=middle><b>�𣠰���</b></td>
    <td class=tablebody1 valign=middle><INPUT name="answer" type=text></td>
    </tr>
    <tr>
    <td class=tablebody1 colspan=2>
    <b>˵����</b>����д����ȷ������𰸡�</td></tr>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="�� ��"></td></tr>
</table>
<input type=hidden value="<%=username%>" name=username>
</form>
<%
		end if
	end if
	rs.close
	set rs=nothing
end sub

sub step2()
	dim answer
	if request("username")="" then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�����������û�����"
		exit sub
	else
		username=replace(request("username"),"'","")
	end if
	if chkpost=false then
   		ErrMsg=ErrMsg+"<Br>"+"<li>���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ���ԡ�"
   		FoundErr=True
		exit sub
	end if
	if request("answer")="" then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��������������𰸡�"
		exit sub
	else
		answer=md5(request("answer"))
	end if
	set rs=conn.execute("select answer,quesion from [user] where username='"&username&"' and answer='"&answer&"'")
	if rs.eof and rs.bof then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�����������𰸲���ȷ�����������롣"
	else
%>
<form action="lostpass.asp?action=step3" method="post"> 
<table cellpadding=3 cellspacing=1 class=tableborder1 align=center>
    <tr>
    <th valign=middle colspan=2 height=25>ȡ�����루���������޸����룩</th></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>�ʣ��⣺</b></td>
    <td class=tablebody1 valign=middle><%=rs(1)%></td>
    </tr>
    <tr>
    <td class=tablebody1 valign=middle><b>�𣠰���</b></td>
    <td class=tablebody1 valign=middle><%=request("answer")%></td>
    </tr>
    <tr>
    <td class=tablebody1 valign=middle><b>�����룺</b></td>
    <td class=tablebody1 valign=middle><input type=password name=password></td>
    </tr>
    <tr>
    <td class=tablebody1 valign=middle><b>ȷ�������룺</b></td>
    <td class=tablebody1 valign=middle><input type=password name=repassword></td>
    </tr>
    <tr>
    <td class=tablebody1 colspan=2>
    <b>˵����</b>
	<%if Forum_Setting(2)>0 then%>
	ϵͳ�����Զ���һ���ʼ�����ע��ʱ��д�����䣬�ڴ��ʼ��е����뼤���ַ�����������뽫��ʽ���á�
	<%else%>
	����д������̳�����룬����ס������д��Ϣ��<%end if%>
	</td></tr>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="�� ��"></td></tr>
</table>
<input type=hidden value="<%=request("answer")%>" name=answer>
<input type=hidden value="<%=username%>" name=username>
</form>
<%
	end if
	rs.close
	set rs=nothing
end sub

sub step3()
	if request("username")="" then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�����������û�����"
		exit sub
	else
		username=replace(request("username"),"'","")
	end if
	if request("answer")="" then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��������������𰸡�"
		exit sub
	else
		answer=md5(request("answer"))
	end if
	if request("password")="" or Len(request("password"))>10 or len(request("password"))<6 then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>����������������(���Ȳ��ܴ���10С��6)��"
		exit sub
	elseif request("repassword")="" then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>���ٴ��������������롣"
		exit sub
	elseif request("password")<>request("repassword") then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��������������ȷ�ϲ�һ������ȷ������д����Ϣ��"
		exit sub
	else
		password=md5(request("password"))
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select userpassword,useremail,quesion,userclass,UserGroupID from [user] where username='"&username&"' and answer='"&answer&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�����������𰸲���ȷ�����������롣"
	else
		if Forum_Setting(2)>0 then
			repassword=request.form("password")
			answer=request.form("answer")
			password=rs("userpassword")
			useremail=rs("useremail")
			call sendusermail()
			if SendMail="OK" then
			sendmsg="ϵͳ�Ѿ�����һ���ʼ�����ע��ʱ��д�����䣬�ڴ��ʼ��е����뼤���ַ�����������뽫��ʽ���á�"
			elseif rs("UserGroupID")<4 then
			sendmsg="����ϵͳ���󣬸��㷢�͵���������δ�ɹ������ĵȼ�Ϊ�������ϼ���Ϊ�˰�ȫ���������ϵ����Ա������롣"
			else
			rs("userpassword")=md5(repassword)
			rs.update
			sendmsg="����ϵͳ���󣬸������͵���������δ�ɹ������Ѿ��޸�����ɹ�����ʹ���������½ϵͳ��"
			end if
		else
			rs("userpassword")=password
			rs.update
		end if
%>
<form action="login.asp" method="post"> 
<table cellpadding=3 cellspacing=1 class=tableborder1 align=center>
    <tr>
    <th valign=middle colspan=2>ȡ�����루���Ĳ����޸ĳɹ���</th></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>�ʣ��⣺</b></td>
    <td class=tablebody1 valign=middle><%=rs(2)%></td>
    </tr>
    <tr>
    <td class=tablebody1 valign=middle><b>�𣠰���</b></td>
    <td class=tablebody1 valign=middle><%=request("answer")%></td>
    </tr>
    <tr>
    <td class=tablebody1 valign=middle><b>�����룺</b></td>
    <td class=tablebody1 valign=middle><%=request("password")%></td>
    </tr>
    <tr>
    <td class=tablebody1 colspan=2>
    <b>˵����</b>
	<%if Forum_Setting(2)>0 then%>
<%=sendmsg%>
	<%else%>
	���ס���������벢ʹ��������<a href=login.asp>��½��̳</a>��
	<%end if%></td></tr>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="�� ��"></td></tr>
</table>
</form>
<%
	end if
	rs.close
	set rs=nothing
end sub

sub main()
%>
<form action="lostpass.asp?action=step1" method="post"> 
<table cellpadding=6 cellspacing=1 align=center class=tableborder1>
    <tr>
    <th valign=middle colspan=2>ȡ������</b>����һ�����û�����</th></tr>
    <tr>
    <td class=tablebody1 valign=middle>�����������û���</td>
    <td class=tablebody1 valign=middle><INPUT name="username" type=text></td>
    </tr>
    <tr>
    <td class=tablebody1 colspan=2>
    <b>˵����</b>������ֻ���޸��������룬���ܶ�ԭ��������޸ģ���ȷ�����Ѿ���д���������⼰�𰸡�</td></tr>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="�� ��"></td></tr>
</table>
</form>
<%
end sub

sub sendusermail()
	on error resume next
	topic="����" & Forum_info(0) & "��������Ϣ"
	mailbody=mailbody &"<style>A:visited {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline	}"
	mailbody=mailbody &"A:link 	  {	text-decoration: none;}"
	mailbody=mailbody &"A:visited {	text-decoration: none;}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none;}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline}"
	mailbody=mailbody &"BODY   {	FONT-FAMILY: ����; FONT-SIZE: 9pt;}"
	mailbody=mailbody &"TD	   {	FONT-FAMILY: ����; FONT-SIZE: 9pt	}</style>"
	mailbody=mailbody &"<TABLE border=0 width='95%' align=center><TBODY><TR>"
	mailbody=mailbody &"<TD valign=middle align=top>"
	mailbody=mailbody &""&htmlencode(username)&"�����ã�<br><br>"
	mailbody=mailbody &"��ӭ��ʹ�ñ���̳�������������ܣ�<b>��������ϣ�������������룬�벻Ҫ�������ļ�������</b>��<br>"
	mailbody=mailbody &"�������������뼤�����ӣ�ֱ�ӽ���������ӵ�ַ�����������������ַ���򿪾Ϳ��Խ����������뼤�<br>"
	mailbody=mailbody &"<a href=http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"lostpass.asp","")&"activepass.asp?username="&htmlencode(username)&"&pass="&password&"&repass="&repassword&"&answer="&answer&">http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"lostpass.asp","")&"activepass.asp?username="&htmlencode(username)&"&pass="&password&"&repass="&repassword&"&answer="&answer&"</a>"
	mailbody=mailbody &"<br><br>"
	mailbody=mailbody &"<center><font color=red>�ٴθ�л��ʹ�ñ�ϵͳ��������һ��������������ϼ�԰��</font>"
	mailbody=mailbody &"</TD></TR></TBODY></TABLE><br><hr width=95% size=1>"
	mailbody=mailbody & Copyright & " &nbsp;&nbsp; " & Version

	select case Forum_Setting(2)
	case 0
	sendmsg="����ϵͳ���󣬸������͵���������δ�ɹ��������ұߵ����ӽ��������뼤�<a href=http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"lostpass.asp","")&"activepass.asp?username="&htmlencode(username)&"&pass="&password&"&repass="&repassword&"><B>��������</B></a>"
	case 1
	call jmail(useremail,topic,mailbody)
	case 2
	call Cdonts(useremail,topic,mailbody)
	case 3
	call aspemail(useremail,topic,mailbody)
	case else
	sendmsg="����ϵͳ���󣬸������͵���������δ�ɹ��������ұߵ����ӽ��������뼤�<a href=http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"lostpass.asp","")&"activepass.asp?username="&htmlencode(username)&"&pass="&password&"&repass="&repassword&"><B>��������</B></a>"
	end select
end sub

%>