<!--#INCLUDE FILE="Conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="char.inc"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
IF request.cookies(Forcast_SN)("KEY")<>"super" THEN
	Show_Err("�Բ������ĺ�̨����Ȩ�޲�����")
	response.end
else
	if request("action")="update" then
		dim typeorder,typemaster,mode,typename,typecontent,typeview
		for i=1 to request.form("typeid").count
			typeid=CheckStr(request.form("typeid")(i))
			mode=CheckStr(request.form("mode")(i))
			typename=CheckStr(request.form("typename")(i))
			typecontent=CheckStr(request.form("typecontent")(i))
			typeorder=CheckStr(request.form("typeorder")(i))
			typemaster=CheckStr(request.form("typemaster")(i))
			typeview=CheckStr(request.form("typeview")(i))
			if CheckStr(request.form("typeorder")(i))="" then
				Show_Err("����д��������<br><br><a href='javascript:history.back()'>����</a>")
				response.end
			end if
			if CheckStr(request.form("typemaster")(i))="" then
				typemaster="��"
			else
				master=split(CheckStr(request.form("typemaster")(i)),"|")
				for k=0 to ubound(master)
					set rs=server.createobject("adodb.recordset")
					sql="Select * from "& db_User_Table &" where oskey='typemaster' and  "& db_User_Name &"='"&trim(master(k))&"'"
					rs.open sql,ConnUser,1,3
					if trim(master(k))<>"��" then
						if rs.eof then
							Show_Err("��������Ա����"& trim(master(k)) &"�û���������ѡ��������Ĺ���Ա��<br><br><a href='javascript:history.back()'>����</a>")
							Response.End
						end if
					else
						typemaster="��"
					end if
					rs.close
					set rs=nothing
				next
			end if
			conn.execute("update "& db_Type_Table &" set typeorder="&typeorder&",typemaster='"&typemaster&"',mode="&mode&",typeview="&typeview&",typename='"&typename&"',typecontent='"&typecontent&"' where typeid="&typeid)
		next
	end if

	if request("action")="add" then
	typecontent=request.form("typecontent")
	typename=trim(request("typename"))
	mode=request("mode")
	typeorder=request("typeorder")
	typeview=request("typeview")
	typemaster=request("typemaster")
	
	If mode="0" Then
		Show_Err("��ѡ��ģ�壡<br><br><a href='javascript:history.back()'>����</a>")
		response.end
	end if
	
	If typeorder="" Then
		Show_Err("����������Ϊ�գ�<br><br><a href='javascript:history.back()'>����</a>")
		response.end
	end if
	If typename="" Then
		Show_Err("�������Ʋ���Ϊ�գ�<br><br><a href='javascript:history.back()'>����</a>")
		response.end
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_Type_Table &""
	rs.open sql,conn,3,3
	do while not rs.eof
		if rs("typename")=typename then
			Show_Err("�Ѿ���������������ƣ�<br><br><a href='javascript:history.back()'>����</a>")
			response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs=nothing
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_Type_Table &""
	rs.open sql,conn,3,3
	do while not rs.eof
		if rs("typeorder")=typeorder then
			Show_Err("�Ѿ������������˳��<br><br><a href='javascript:history.back()'>����</a>")
			response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs=nothing
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from "& db_Type_Table &""
	rs.open sql,conn,3,3
	rs.addnew
	rs("typename")=typename
	rs("mode")=mode
	rs("typecontent")=typecontent
	rs("typeorder")=typeorder
	rs("typeview")=typeview
	if typemaster="" then
		rs("typemaster")="��"
	else
		rs("typemaster")=typemaster
	end if
	rs.update
	rs.close
	set rs=nothing
	end if
	
	conn.close
	set conn=nothing
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=TypeManage.asp"">"
	Show_Message("<p align=center><font color=red>��ϲ��!���óɹ�!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
	response.end
end if%>