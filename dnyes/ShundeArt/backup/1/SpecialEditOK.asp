<!--#INCLUDE FILE="Conn.asp" -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="chkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	dim SpecialName
	dim SpecialID
	dim Specialzs
	dim name
	name=request.form("name")
	Specialzs=request.form("specialzs")
	SpecialID=request("SpecialID")
	SpecialName=trim(request("SpecialName"))
	
	if SpecialName="" then
		Show_Err("ר�����Ʋ���Ϊ�գ�<br><br>������<a href='javascript:history.back()'>��ȥ����</a>������")
		response.end
	end if
	
	'�޸�ר���
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_Special_Table &" where SpecialID="&SpecialID
	rs.open sql,conn,3,3
	rs("SpecialName")=SpecialName
	rs("Specialzs")=Specialzs
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=SpecialManage.asp"">"
	Show_Message("<p align=center><font color=red>��ϲ��!���޸�ר�⡰"&name&"��Ϊ��"&SpecialName&"���ɹ�!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
	response.end
end if%>