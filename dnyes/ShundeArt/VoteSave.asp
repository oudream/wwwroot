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
		ID=request.QueryString("id")
		Title=trim(request.form("Title"))
		act=request("act")
		if DateAndTime="" then DateAndTime=now()
		Content=trim(request.form("Content"))
		founerr=false
		if Title="" then
			Show_Err("�Բ���ͶƱ���ⲻ��Ϊ�գ�<br><br><a href='javascript:history.back()'>��ȥ����</a>")
			response.end
		end if
		if founderr<>true then
			set rs=server.createobject("adodb.recordset")
			if act="edit" then
				sql="select * from "& db_Vote_Table &" where ID="&ID
			elseif act="add" then
					sql="select * from "& db_Vote_Table &""
				else
					Show_Err("<li>������������ϵ����Ա</li><br><br><a href='javascript:history.back()'>��ȥ����</a>")
					Response.End 
			end if
			rs.open sql,conn,1,3
			if act="add" or act="edit" then
				if act="edit" then
					if rs.eof then
						Show_Err("<li>������������ϵ����Ա</li><br><br><a href='javascript:history.back()'>��ȥ����</a>")
						Response.End 
					else
						for i=1 to 8
							if request("select"&i)="" then
								rs("select"&i)=""
							end if 
						next
					end if
				end if
				if act="add" then
					rs.addnew
					rs("Title")=Title
					for i=1 to 8
						if request("select"&i)<>"" then
							rs("select"&i)=request("select"&i)
							if request("answer"&i)="" then
								rs("answer"&i)=0
							else
								rs("answer"&i)=request("answer"&i)
							end if
						end if
					next
					rs("dateandtime")=now()
					rs("IsChecked")=0
					rs.update
				end if
			end if
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Redirect "VoteManage.asp"
		end if
	else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if%>