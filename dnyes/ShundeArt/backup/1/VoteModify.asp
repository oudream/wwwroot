<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
IF not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster" or request.cookies(Forcast_SN)("ManageKEY")="check" or request.cookies(Forcast_SN)("ManageKEY")="typemaster") THEN
	Show_Err("�Բ������Ĺ���Ȩ�޲�����")
	response.end
else
	if votemana="1" or (request.cookies(Forcast_SN)("ManagePurview")="99999" and request.cookies(Forcast_SN)("purview")="99999") then
		id=request.QueryString("id")
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_Vote_Table &" where id="&id
		rs.open sql,conn,1,1
		if rs.eof then
			Show_Err("<li>������������ϵ����Ա<br><br>������<a href='javascript:history.back()'>��ȥ����</a>������")
			Response.End 
		else
			Title=rs("Title")
			DateAndTime=rs("DateAndTime")
		end if
		%>
<head><head>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ͶƱ����</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>" align=center>
<form method="POST" action="VoteSave.asp?id=<%=id%>">
<tr class="TDtop">
	<td width="100%" height="25" colspan=2 align=center>�� <strong>�� �� �� վ Ͷ Ʊ</strong> ��</td>
</tr>
<tr>
	<td colspan=2 align=center></td>
</tr>
<tr>
	<td width="40%" align="right">���⣺</td>
	<td width="60%"><input type="text" name="Title" value="<%=Title%>" size="20" style="font-family: ����; font-size: 9pt"></td>
</tr>
		<%for i=1 to 8%>
<tr>
	<td align="right">ѡ��<%=i%>��</td>
	<td>
		<input type="text" name="select<%=i%>" value="<%=rs("select"&i)%>" size="20" style="font-family: ����; font-size: 9pt"> Ʊ����<input type="text" name="answer<%=i%>" value="<%=rs("answer"&i)%>" size="5" style="font-family: ����; font-size: 9pt">
	</td>
</tr>
		<%next%>
<tr>
	<td align="right">����ʱ�䣺</td>
	<td>
		<input type="text" name="dat" value="<%=DateAndTime%>" disabled size="20" style="font-family: ����; font-size: 9pt">
	</td>
</tr>
<tr>
	<td colspan=2 align=center></td>
</tr>
<tr class="TDtop">
	<td colspan=2 align=center>
		<input type="hidden" value="edit" name="act">
		<input type="button" value=" �� �� " onclick="javascript:history.go(-1)" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp; 
		<input type="submit" value=" �� �� " name="cmdok" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp; 
		<input type="reset" value=" �� �� "  name="cmdcancel" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr>
	<td colspan=2 align=center></td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if%>