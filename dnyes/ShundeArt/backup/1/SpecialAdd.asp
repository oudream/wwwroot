<!--#INCLUDE FILE="Conn.asp" -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="chkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<!--#include file="Function.asp"-->

<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	
	dim Specialname
	specialzs=CheckStr(trim(request.form("specialzs")))
	Specialname=CheckStr(trim(request("Specialname")))
	
	If Specialname="" Then
		Show_Err("ר�����Ʋ���Ϊ�գ�<br><br>������<a href='javascript:history.back()'>��ȥ����</a>������")
		response.end
	end if
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_Special_Table &""
	rs.open sql,conn,3,3
	do while not rs.eof
		if rs("Specialname")=Specialname then
			Show_Err("�Ѿ��������ר�⣡<br><br>������<a href='javascript:history.back()'>����������д</a>������")
			response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs=nothing

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from "& db_Special_Table &""
	rs.open sql,conn,3,3
	rs.addnew
	rs("Specialname")=Specialname
	rs("Specialzs")=Specialzs
	rs("Specialmaster")=request.cookies(Forcast_SN)("ManageUserName")
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=SpecialManage.asp"">"
	Show_Message("<p align=center><font color=red>��ϲ��!�����ר�⡰"&Specialname&"���ɹ�!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
	response.end
end if
%>