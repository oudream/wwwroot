<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp"-->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
IF request.cookies(Forcast_SN)("ManageKEY")<>"super" then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲�����")
	response.end
else
	ID = Request("ID")
	set rs=server.createobject("adodb.recordset")
	sql="Select * from "& db_dep_Table &" where ID="&id
	rs.open sql,conn,1,3
	depname=rs("depname")
	rs.close
	set rs=nothing
	
	button_value=trim(Request.Form("alert_button"))
	if button_value="ȷ��" then
		conn.execute("delete from "& db_dep_Table &" where ID=" &ID)
		conn.close
		set conn=nothing
	end if
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=DepManage.asp"">"
	Show_Message("<p align=center><font color=red>��λɾ���ɹ�!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
	response.end
end if%>