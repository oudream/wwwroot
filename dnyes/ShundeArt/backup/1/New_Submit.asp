<%@ Language=VBScript %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp" -->
<%
if not((request.cookies(Forcast_SN)("ManageKEY")="super" and request.cookies(Forcast_SN)("ManagePurview")="99999") or (request.cookies(Forcast_SN)("KEY")="super" and request.cookies(Forcast_SN)("purview")="99999")) then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	if init="1" then
		if (request.form("action")="del") then
			if request.form("selectdel")<>"" then
				For i = 1 To Request.Form("selectdel").Count
					SelectTable=Request.Form("selectdel")(i)
					select case SelectTable
						'��ʼ�����ű�
						case db_Dep_Table conn.execute "delete from "& db_Dep_Table
						'��ʼ��������
						case db_Type_Table conn.execute "delete from "& db_Type_Table
						'��ʼ��������
						case db_BigClass_Table conn.execute "delete from "& db_BigClass_Table
						'��ʼ��С����
						case db_SmallClass_Table conn.execute "delete from "& db_SmallClass_Table
						'��ʼ��ר���
						case db_Special_Table conn.execute "delete from "& db_Special_Table
						'��ʼ�����±�
						case db_News_Table conn.execute "delete from "& db_News_Table
						'��ʼ����ʱ��
						case db_NewsFile_Table conn.execute "delete from "& db_NewsFile_Table
						'��ʼ���ϴ���ʱ��
						case db_UploadPic_Table conn.execute "delete from "& db_UploadPic_Table
						'��ʼ��ͶƱ��
						case db_Vote_Table conn.execute "delete from "& db_Vote_Table
						'��ʼ���������ӱ�
						case db_Link_Table conn.execute "delete from "& db_Link_Table
						'��ʼ�������
						case db_Board_Table conn.execute "delete from "& db_Board_Table
						'��ʼ��������
						case db_Attach_Table conn.execute "delete from "& db_Attach_Table
						'��ʼ�����۱�
						case (db_Review_Table&"1") conn.execute "delete from "& db_Review_Table &" where NewsId>0"
						'��ʼ�����Ա�
						case (db_Review_Table&"0") conn.execute "delete from "& db_Review_Table &" where NewsId<1"
						'��ʼ���û���
						case db_User_Table ConnUser.execute "delete from "& db_User_Table &" where purview<>99999"
						'��ʼ������Ա��
						case db_Manage_Table conn.execute "delete from "& db_Manage_Table &" where purview<>99999"
					end select
				Next
				ConnUser.close
				set ConnUser=nothing
				conn.close
				set conn=nothing
				response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=New.asp"">"
				Show_Message("<p align=center><font color=red>ϵͳ��"& (request.form("selectdel")) &"��ʼ�����!<br><br>"&freetime&"���Ӻ󷵻س�ʼҳ!</font>")
				response.end
			else
				Show_Err("�Բ�������û��ѡ���κβ�����<br><br>������<a href='javascript:history.back()'>��ȥ����</a>������")
				response.end
			end if
		else
			SpecialID=Request.Form("SpecialID")
			button_value=trim(Request.Form("alert_button"))
			if button_value="ȷ��" then
				conn.execute("delete from "& db_BigClass_Table &"")
				conn.execute("delete from "& db_SmallClass_Table &"")
				conn.execute("delete from "& db_Special_Table &"")
				conn.execute("delete from "& db_News_Table &"")
				conn.execute("delete from "& db_Review_Table &"")
				conn.execute("delete from "& db_Vote_Table &"")
				conn.execute("delete from "& db_Link_Table &"")
				conn.execute("delete from "& db_Type_Table &"")
				conn.execute("delete from "& db_Dep_Table &"")
				conn.execute("delete from "& db_Board_Table &"")
				conn.execute("delete from "& db_Attach_Table &"")
				conn.execute("delete from "& db_NewsFile_Table &"")
				conn.execute("delete from "& db_UploadPic_Table &"")
				if UserTableType = "FeiTium" then		'��������
					ConnUser.execute("delete from "& db_User_Table &" where purview<>99999")
				end if
				response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=Admin_login.asp"">"
				Show_Message("<p align=center><font color=red>ϵͳ�Ѿ�����ʼ��!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
				response.end
			else
				Response.Redirect "Admin_login.asp"
			end if
		end if
	else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if%>