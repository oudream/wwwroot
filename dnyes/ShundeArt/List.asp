<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp" -->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	dim sql
	dim rst
	set rst=server.createobject("adodb.recordset")
	sql="select * from "& db_User_Table &" where "&Db_User_ID&"="&request("id")
	rst.open sql,ConnUser,1,1
	if UserTableType = "Dvbbs" then
		photowidth=rst(db_User_FaceWidth)		''ȡ��̳�趨��ͼƬ���
		photoheight=rst(db_User_FaceHeight)		''ȡ��̳�趨��ͼƬ�߶�
	end if
	%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - <%=rst("fullname")%> ����ϸ����</title>
</head>
<body topmargin="0">
<!--#include file="Admin_Top.asp"-->
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr> 
	<td colspan="2">
		<div align="right">
			��ţ�<%=rst(db_User_Id)%>&nbsp;&nbsp; 
	<%if rst(db_User_RegDate)<>"" then%>
			ע�����ڣ�<%=rst(db_User_RegDate)%> 
	<%end if%>
		</div>
	</td>
</tr>
<tr> 
	<td width="614"> 
	<%if rst("fullname")<>"" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp; ����<%=rst("fullname")%><br> <br> 
	<%end if%>
	<%if db_Sex_Select = "FeiTium" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;&nbsp;��<%=rst(db_User_sex)%> 
	<%else%>
		<%if db_Sex_Select = "Number" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;&nbsp;�� 
			<%if rst(db_User_Sex)="0" then%>
		Ů
			<%else%>
		��
			<%end if%>
		<%end if%>
	<%end if%>
	<br> <br> 
	<%if db_Birthday_Select = "FeiTium" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp; �գ�<%=rst(db_User_Birthyear)%>��<%=rst(db_User_Birthmonth)%>��<%=rst(db_User_Birthday)%>�� 
	<%else%>
		<%if db_Birthday_Select = "Text" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp; �գ�<%=year(rst(db_User_birthday))%>��<%=month(rst(db_User_birthday))%>��<%=day(rst(db_User_birthday))%>�� 
		<%end if%>
	<%end if%>
	<br> <br> 
	<%if rst(db_User_Email)<>"" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;�����ʼ���<a href=mailto:<%=rst(db_User_Email)%>><%=rst(db_User_Email)%></a><br> 
	<br> 
	<%end if%>
	<%if rst("tel")<>"" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;��ϵ�绰��<%=rst("tel")%><br> <br> 
	<%end if%>
		&nbsp;&nbsp;&nbsp;&nbsp;Ȩ&nbsp;&nbsp;&nbsp; �ޣ�
	<%oskey1=rst("oskey")
	select case oskey1
		case "super"
			if rst("purview")="99999" then
				response.write "�����û�"
			else
				response.write "ϵͳ�û�"
			end if
		case "typemaster" response.write "�����û�"
		case "bigmaster" response.write "�����û�"
		case "smallmaster" response.write "С���û�"
		case "selfreg" response.write "ע���û�"
		case "checkf" response.write "����û�"
	end select%>
		<br> <br> &nbsp;&nbsp;&nbsp;&nbsp;������λ��<%=rst("depname")%><br> <br>
		&nbsp;&nbsp;&nbsp;&nbsp;���ڲ��ţ�<%=rst("deptype")%><br> <br>&nbsp;&nbsp;&nbsp;&nbsp;��¼������<%=rst(db_User_LoginTimes)%><br> 
		<br> &nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;��&nbsp;����<%=rst("number")%><br> <br> 
		&nbsp;&nbsp;&nbsp;&nbsp;�����ַ��<%=rst(db_User_LastLoginIP)%><br> <br> &nbsp;&nbsp;&nbsp;&nbsp;�����¼��<%=rst(db_User_LastLoginTime)%><br> 
		<br> &nbsp;&nbsp;&nbsp;&nbsp;���ҽ��ܣ�<%=rst("content")%> </td>
		<td width="137"><div align="center"> 
		<%if rst(db_User_Face) <> "" then
			if UserTableType = "FeiTium" then%>
		<img src="<%=rst(db_User_Face)%>" border="0" width="120"> 
			<%else
				if UserTableType = "Dvbbs" then%>
					<img src="<%=Bbspath%><%=rst(db_User_Face)%>" border="0" width="<%=rst(db_User_FaceWidth)%>" height="<%=rst(db_User_FaceHeight)%>"> 
					<%''��ʾ�û�ͷ�񣬼�'bbs/'ǰ׺·��,ʹͼ��ϵͳֱ����ʾ������̳ͷ��%>
				<%end if
			end if
		else%>
			<img src="images/nopic.gif" border="0"> 
		<%end if%>
		<br>
		<br>
		�û�ͷ��</div></td>
</tr>
<tr> 
	<td colspan="2">
		<div align="right">[<a href="UserEdit.asp?id=<%=rst(db_User_Id)%>">�޸�</a>] 
		<%if trim(session("username"))<>trim(rst(db_User_Name)) then%>
            <% if trim(request.cookies(Forcast_SN)("username"))=trim(rst(db_User_Name)) then %>
			    <%response.write "[----]"%>
			<%else%>
			    [<a href="UserDel.asp?id=<%=rst(db_User_Id)%>&amp;name=del">ɾ��</a>] 
		    <%end if%>
		<%end if%>
		[<a href="javascript:history.go(-1)">����</a>]</div></td>
</tr>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%
	rst.close
	set rst=nothing
	conn.close
	set conn=nothing
end if%>