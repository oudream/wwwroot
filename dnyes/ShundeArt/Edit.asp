<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<%
IF request.cookies(Forcast_SN)("KEY")="" THEN
	Show_Err("�Բ�����Ȩ�޲�����")
	response.end
else
	usernamecookie=CheckStr(request.cookies(Forcast_SN)("UserName"))
	passwdcookie=replace(trim(Request.cookies(Forcast_SN)("passwd")),"'","''")
	KEYcookie=replace(trim(request.cookies(Forcast_SN)("KEY")),"'","''")
	%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ���������޸�</title>
<script LANGUAGE="javascript">
<!--
function FrmAddLink_onsubmit() {
 if (document.FrmAddLink.passwd.value=="")
	{
	  alert("�Բ��������������룡")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length < 4)
	{
	  alert("Ϊ�˰�ȫ����������Ӧ�ó�һ�㣡")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length > 16)
	{
	  alert("��������̫���˰ɣ�")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.username.value==document.FrmAddLink.passwd.value)
	{
	  alert("Ϊ�˰�ȫ���û��������벻Ӧ����ͬ��")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value=="")
	{
	  alert("�Բ�������������֤���룡")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value !== document.FrmAddLink.passwd.value)
	{
	  alert("�Բ�����������������벻һ�£�")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
else if (document.FrmAddLink.question.value=="")
	{
	  alert("�Բ�������������ʾ���⣡")
	  document.FrmAddLink.question.focus()
	  return false
	 }
else if (document.FrmAddLink.answer.value=="")
	{
	  alert("�Բ���������������𰸣�")
	  document.FrmAddLink.answer.focus()
	  return false
	 }
else if (document.FrmAddLink.question.value==document.FrmAddLink.answer.value)
	{
	  alert("Ϊ�˰�ȫ����ʾ����������𰸲�Ӧ����ͬ��")
	  document.FrmAddLink.answer.focus()
	  return false
	 }
else if (document.FrmAddLink.fullname.value=="")
	{
	  alert("�Բ���������������ʵ������")
	  document.FrmAddLink.fullname.focus()
	  return false
	 }
else if (document.FrmAddLink.depid.value=="")
	{
	  alert("�Բ������������Ĺ�����λ��")
	  document.FrmAddLink.depid.focus()
	  return false
	 }
else if (document.FrmAddLink.sex.value=="")
	{
	  alert("�Բ��������������Ա�")
	  document.FrmAddLink.sex.focus()
	  return false
	 }
else if (document.FrmAddLink.tel.value=="")
	{
	  alert("�Բ���������������ϵ�绰��")
	  document.FrmAddLink.tel.focus()
	  return false
	 }
else if (document.FrmAddLink.email.value=="")
	{
	  alert("�Բ������������ĵ����ʼ���")
	  document.FrmAddLink.email.focus()
	  return false
	 }
else if (document.FrmAddLink.email.value.indexOf("@",0)== -1||document.FrmAddLink.email.value.indexOf(".",0)==-1)
	  {
	  alert("�Բ���������ĵ����ʼ�����")
	  document.FrmAddLink.email.focus()
	  return false
	 }
}
//-->
</script>
<!--����ѡ�����ڴ���ʼ-->
<SCRIPT language=JavaScript src="inc/User_Info_Modify.js"></SCRIPT>
<!--����ѡ�����ڴ������-->
</head>
<body topmargin="0">
<!--#include file="Admin_Top.asp"-->
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="POST" name="FrmAddLink" LANGUAGE="javascript" onsubmit="return FrmAddLink_onsubmit()"  action="Save.asp" >
	<%dim logins
	dim sql
	dim rs
	sql="select * from "& db_User_Table &" where "& db_User_Name &"='"&usernamecookie&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,ConnUser,1,1
	logins=rs(db_User_LoginTimes)
	if UserTableType = "Dvbbs" then
		photowidth=rs(db_User_FaceWidth)		''ȡ��̳�趨��ͼƬ���
		photoheight=rs(db_User_FaceHeight)		''ȡ��̳�趨��ͼƬ�߶�
	end if
	%> 
<tr>
	<td width="100%" class="TDtop" colspan="2" height=25>
		<p align="center" >
			<%if logins=1 then%>
			<font color=red>�������״�ʹ�ñ�ϵͳ��Ϊ�˰�ȫ���������޸����ĸ������Ϻ���ִ�й��������</font><br>
			<%end if%>
			<%=request.cookies(Forcast_SN)("fullname")%>�޸�����
		</p>
	</td>
</tr>
<tr>
	<td width="46%"><p align="right">�� �� ����</td>
	<td width="54%">
		<input name="username" size="30" value="<%=rs(db_User_Name)%>" readonly class="smallInput" maxlength="30" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr>
	<td><p align="right">��&nbsp;&nbsp;&nbsp; �룺</td>
	<td>
		<input name="passwd" type="password" size="30" value="<%=rs(db_User_Password)%>" class="smallInput" maxlength="50" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" <%if request.cookies(Forcast_SN)("key")="smallmaster" and inputmodpwd="0" then%>readonly<%end if%>><%if request.cookies(Forcast_SN)("key")="smallmaster" and inputmodpwd="0" then%>���û��������޸�����<%end if%>
	</td>
</tr>
<tr>
	<td><p align="right">��֤���룺</td>
	<td>
		<input name="passwd2" type="password" size="30" value="<%=rs(db_User_Password)%>" class="smallInput" maxlength="50" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" <%if request.cookies(Forcast_SN)("key")="smallmaster" and inputmodpwd="0" then%>readonly<%end if%>><%if request.cookies(Forcast_SN)("key")="smallmaster" and inputmodpwd="0" then%>���û��������޸�����<%end if%>
	</td>
</tr>
<tr>
	<td><p align="right">��ʾ���⣺</td>
	<td>
		<input name="question" type="text" size="30" value="<%=rs(db_User_Question)%>" class="smallInput"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr>
	<td><p align="right">����𰸣�</td>
	<td>
		<input name="answer" type="text" size="30" value="<%=rs(db_User_Answer)%>" class="smallInput" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr>
	<td><p align="right">��ʵ������</td>
	<td>
		<input name="fullname" type="text" size="30" value="<%=rs("fullname")%>" class="smallInput" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr>
	<td><p align="right">������λ��</td>
	<td>
	<%
	set rstype=createobject("adodb.recordset")
	sql="select * from "& db_dep_Table &" order by id"
	rstype.Open sql,conn,1,3
	%>
		<select name="depid" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
			<option value="">��ѡ������λ</option>
	<%do while not rstype.EOF%>
			<option value="<%=rstype("id")%>" <%if rs("depid")=cint(rstype("id")) then%> selected <%end if%>><%=rstype("depname")%></option>
		<%
		rstype.MoveNext
	loop
	set rstype=nothing
	%>
		</select>
	</td>
</tr>
<tr>
	<td><p align="right">��&nbsp;&nbsp;&nbsp; ��</td>
	<td> 
	<%if db_Sex_Select= "FeiTium"	then %>
		<select size="1" name="sex" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" >
			<option <%if rs(db_User_sex)="����" then%>selected<%end if%> value="����">����</option>
			<option <%if rs(db_User_sex)="Ůʿ" then%>selected<%end if%> value="Ůʿ">Ůʿ</option>
			<option <%if rs(db_User_sex)="����" then%>selected<%end if%> value="����">����</option>
		</select>
	<%else%>
		<%if  db_Sex_Select="Number" then%>
		<select size="1" name="sex" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" >
			<option <%if rs(db_User_Sex)=1 then%>selected<%end if%> value="1">����</option>
			<option <%if rs(db_User_Sex)=0 then%>selected<%end if%> value="0">Ůʿ</option>
		</select>
		<%end if%>
	<%end if%>
	</td>
</tr>
<tr>
	<td><p align="right">��&nbsp;&nbsp;&nbsp; �գ�</td>
	<td > 
		<div align="left"> 
	<%if db_Birthday_Select= "FeiTium" then %>
			<select size="1" name="birthyear" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1950 to 2000%>
				<option <%if rs(db_User_birthyear)=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
			</select>��
			<select size="1" name="birthmonth" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 12%>
				<option <%if rs(db_User_birthmonth)=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
			</select>��
			<select size="1" name="birthday" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 31%>
				<option <%if rs(db_User_birthday)=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
			</select>��
	<%else
		if db_Birthday_Select="Text" then %>
			<INPUT  name=birthday onfocus="show_cele_date(birthday,'','',birthday)" value="<%=rs(db_User_birthday)%>" class="smallInput" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></div>
		<%end if%>
	<%end if%>
	</td>
</tr>
<tr>
	<td><p align="right">��ϵ�绰��</td>
	<td valign="middle"> 
		<input name="tel" size="30" class="smallInput" value="<%=rs("tel")%>" maxlength="100" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" >
	</td>
</tr>
<tr>
	<td><p align="right">�����ʼ���</td>
	<td>
		<input name="email" type="text" size="30" value="<%=rs(db_User_Email)%>" class="smallInput"   style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr>
	<td><p align="right">������Ƭ��</td>
	<td>
		<iframe name="ad" frameborder=0 width=100% height=20 scrolling=no src='UploadFace.asp'></iframe>
		<input name="photo" type="text" size="35" value="<%=rs(db_User_Face)%>" class="smallInput" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="������Ƭ���������ϴ��Լ�����Ƭ��Ҳ����ֱ����д����������Ƭ�ĵ�ַ��">
	<%if UserTableType = "FeiTium" then
	else%>
		<p align=center>
		<%if UserTableType = "Dvbbs" then
			if rs(db_User_Face)<>"" then%>
				<img src="<%=BbsPath%><%=rs(db_User_Face)%>" border="0" width="<%=photowidth%>" height="<%=photoheight%>"> 
				<%''��ʾ�û�ͷ�񣬼�'bbs/'ǰ׺·��,ʹͼ��ϵͳֱ����ʾ������̳ͷ��%>
			<%else%>
				<img src="images/nopic.gif" border="0">
			<%end if
		end if%>
		<br><br>
		�û�ͷ��
		</p>
	<%end if%>
	</td>
</tr>
<tr>
	<td><p align="right">���ҽ��ܣ�</td>
	<td><textarea rows="10" name="content" cols="30"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"><%=rs("content")%></textarea></td>
</tr>
<tr>
	<td width="100%" colspan="2">
		<p align="center">
			<input type="submit" value="�� ��" name="cmdok" class="buttonface" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp; 
			<input type="reset" value="�� ��" name="cmdcancel" class="buttonface" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp; 
			<input type="button" value="�� ��" onclick="javascript:history.go(-1)" class="buttonface" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		</P>
	</td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	ConnUser.close
	set ConnUser=nothing
end if%>