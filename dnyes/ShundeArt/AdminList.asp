<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="md5.asp"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp"-->
<!--#include file="ChkURL.asp"-->

<%
IF not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
else
	dim sql
	dim rst
	set rst=server.createobject("adodb.recordset")
	sql="select * from "& db_Manage_Table &" where "& db_ManageUser_ID &"="&request("id")
	rst.open sql,Conn,1,1
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
		<div align="right">��ţ�<%=rst(db_ManageUser_Id)%>&nbsp;&nbsp;
	<%if rst(db_ManageUser_RegDate)<>"" then%>
			ע�����ڣ�<%=rst(db_ManageUser_RegDate)%> 
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
		&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;&nbsp;��<%=rst(db_ManageUser_sex)%> 
	<%else%>
		<%if db_Sex_Select = "Number" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;&nbsp;��<%if rst(db_ManageUser_Sex)="0" then%>Ů<%else%>��<%end if%>
		<%end if%>
	<%end if%>
	<br> <br> 
	<%if db_Birthday_Select = "FeiTium" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp; �գ�<%=rst(db_ManageUser_Birthyear)%>��<%=rst(db_ManageUser_Birthmonth)%>��<%=rst(db_ManageUser_Birthday)%>�� 
	<%else%>
		<%if db_Birthday_Select = "Text" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp; �գ�<%=year(rst(db_ManageUser_birthday))%>��<%=month(rst(db_ManageUser_birthday))%>��<%=day(rst(db_ManageUser_birthday))%>�� 
		<%end if%>
	<%end if%>
	<br> <br> 
	<%if rst(db_ManageUser_Email)<>"" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;�����ʼ���<a href=mailto:<%=rst(db_ManageUser_Email)%>><%=rst(db_ManageUser_Email)%></a><br> 
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
				response.write "��������Ա"
			else
				response.write "ϵͳ����Ա"
			end if
		case "typemaster" response.write "��������Ա"
		case "bigmaster" response.write "�������Ա"
		case "smallmaster" response.write "С�����Ա"
		case "selfreg" response.write "ע���û�"
		case "checkf" response.write "��˹���Ա"
	end select%>
		<br><br>&nbsp;&nbsp;&nbsp;&nbsp;������λ��<%=rst("depname")%>
		<br><br>&nbsp;&nbsp;&nbsp;&nbsp;���ڲ��ţ�<%=rst("deptype")%>
		<br><br>&nbsp;&nbsp;&nbsp;&nbsp;��¼������<%=rst(db_ManageUser_LoginTimes)%>
		<br><br>&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;��&nbsp;����<%=rst("number")%>
		<br><br>&nbsp;&nbsp;&nbsp;&nbsp;�����ַ��<%=rst(db_ManageUser_LastLoginIP)%>
		<br><br>&nbsp;&nbsp;&nbsp;&nbsp;�����¼��<%=rst(db_ManageUser_LastLoginTime)%>
		<br><br>&nbsp;&nbsp;&nbsp;&nbsp;���ҽ��ܣ�<%=rst("content")%>
	</td>
	<td width="137"><div align="center"> 
		<%if rst(db_User_Face) <> "" then%>
			<%if UserTableType = "FeiTium" then%>
		         <img src="<%=rst(db_User_Face)%>" border="0" width="120">
			<%else%>
				 <%if UserTableType = "Dvbbs" then%>
					<img src="<%=Bbspath%><%=rst(db_User_Face)%>" border="0" width="<%=rst(db_User_FaceWidth)%>" height="<%=rst(db_User_FaceHeight)%>"> 
					<%''��ʾ�û�ͷ�񣬼�'bbs/'ǰ׺·��,ʹͼ��ϵͳֱ����ʾ������̳ͷ��%>
				 <%end if%>
			<%end if%>
		<%else%>
			<img src="images/nopic.gif" border="0"> 
		<%end if%>
		<br>
		<br>
		�û�ͷ��</div>
    </td>
</tr>
<tr> 
	<td colspan="2">
		<div align="right">[<a href="AdminEdit.asp?id=<%=rst(db_ManageUser_Id)%>">�޸�</a>] 
		<%if trim(session("username"))<>trim(rst(db_ManageUser_Name)) then%>
            <% if trim(request.cookies(Forcast_SN)("username"))=trim(rst(db_ManageUser_Name)) then %>
			    <%response.write "[----]"%>
			<%else%>
			    [<a href="AdminDel.asp?id=<%=rst(db_ManageUser_Id)%>&amp;name=del">ɾ��</a>] 
		    <%end if%>
		<%end if%>
			[<a href="javascript:history.go(-1)">����</a>]
		</div>
	</td>
</tr>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
<%
END IF
rst.close
set rst=nothing
conn.close
set conn=nothing
%>