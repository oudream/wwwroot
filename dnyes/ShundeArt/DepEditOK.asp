<!--#INCLUDE FILE="Conn.asp" -->
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
	function changechr(str) 
		changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," "," ") 
		changechr=replace(changechr,"'","&acute;")
		changechr=replace(changechr,mid(" "" ",2,1),"&quot;")
	end function

	dim depname
	dim deptype
	dim ID
	ID=request("ID")
	deptype=changechr(trim(request.form("deptype")))
	name=request.form("name")
	depname=changechr(trim(request("depname")))
	if depname="" or deptype="" then
		Show_Err("��λ�������Ʋ���Ϊ��<br><br><a href=javascript:history.go(-1)>����������д</a>")
		response.end
	else
		'�޸Ĳ��ſ�
		dim depid
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_dep_Table &" where ID="&ID
		rs.open sql,conn,3,3
		depid=rs("id")
		rs("depName")=depname
		rs("deptype")=deptype
		rs.update
		rs.close
		set rs=nothing
	
		'�޸��û���
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_User_Table &" where depid="&depid
		rs.open sql,ConnUser,3,3
		do while not rs.eof
			rs("depName")=depname
			rs.update
			rs.movenext
		loop
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=DepManage.asp"">"
		Show_Message("<p align=center><font color=red>��λ�����޸ĳɹ�!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
		response.end
	end if
end if%>