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
		Checked=request.form("Checked")
		if Checked="" then
			Show_Err("�Բ���������ѡ��һ��ͶƱ�<br><br><a href='javascript:history.back()'>��ȥ����</a>")
			Response.End 
		end if
		set rs=server.createobject("adodb.recordset")
		sql="update "& db_Vote_Table &" set IsChecked=0"
		rs.open sql,conn,1,3
		rs.close
	
		sql="Select * from "& db_Vote_Table &" where ID="&Checked
		rs.open sql,conn,1,3
		if not rs.EOF then
			do while not rs.EOF 
				rs("isChecked")=1
			rs.MoveNext 
			loop
		end if
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.redirect "VoteManage.asp"
	else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if
%>