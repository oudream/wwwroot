<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	if request.cookies(Forcast_SN)("ManageKEY")<>"bigmaster" then
		response.redirect "CheckNews1.asp"
		response.end
	end if
	%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - �������</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<br>
<div align="center">
<table border="1" width="100%" bgcolor="<%=border%>" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>">
<form method="POST" action="checknews1.asp">
<tr align="center">
	<td height="55" bgcolor="<%=m_top%>"><b>�� �� �� ��</b></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" height="197">
	<%
	set rs=server.CreateObject ("ADODB.RecordSet")
	rs.Source="select * from "& db_BigClass_Table &" order by BigClassorder"
	rs.Open rs.source,conn,1,1
	%>
		<p>��ѡ�����´��ࣺ
		<select size="1" name="BigClassid" onMouseOver="window.status='������ѡ��Ҫ��˵����´���';return true;" onMouseOut="window.status='';return true;">
	<%if rs.EOF then %>
			<option value="">�����κ����</option>
	<%else
		if request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="check" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" then
			while not rs.EOF
				%>
			<option value='<%=rs("BigClassid")%>'><%=trim(rs("BigClassName"))%></option>
				<%
				rs.MoveNext
			wend
		else
			while not rs.EOF
				if instr(rs("BigClassMaster"),Session("UserName"))>0 then
					%>
			<option value='<%=rs("BigClassid")%>'><%=trim(rs("BigClassName"))%></option>
					<%
				end if
				rs.MoveNext
			wend
		end if
	end if
	%>
		</select>
		</p>
		<p></p>
		<p>&nbsp;</p>
	</td>
</tr>
<tr>
	<td align="center" height="55" bgcolor="<%=m_top%>">
		<input type="submit" value="  ��һ��>>  " name="B1" onMouseOver="window.status='�������ť������һ��';return true;" onMouseOut="window.status='';return true;" title='�������ť������һ��'>
	</td>
</tr>
</form>
</table>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
<%
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if%>