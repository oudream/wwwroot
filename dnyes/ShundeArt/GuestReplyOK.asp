<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkURL.asp"-->
<%if request.cookies(Forcast_SN)("key")="super" then
	usernamecookie=request.cookies(Forcast_SN)("UserName")
	passwdcookie=CheckStr(trim(Request.cookies(Forcast_SN)("passwd")))
	KEYcookie=CheckStr(trim(request.cookies(Forcast_SN)("KEY")))
	if usernamecookie="" or passwdcookie="" then
		response.write "<script>alert('����Ȩ�ظ������ԣ�');history.go(-1);</Script>"
		response.end
	else
		'�ж��û��ĺϷ���
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_User_Table &" where "& db_User_Name &"='"&usernamecookie&"'"
		rs.open sql,ConnUser,1,1
		if rs.eof and rs.bof then
			response.write "<script>alert('����Ȩ�ظ������ԣ�');history.go(-1);</Script>"
			response.end
		end if
		IF passwdcookie<>rs(db_User_Password) THEN
			response.write "<script>alert('����Ȩ�ظ������ԣ�');history.go(-1);</Script>"
			response.end
		END IF
		'�����ж��û�����ʵ�������û������Ƕ�Ӧ���ж�
		if KEYcookie<>rs("OSKEY") then
			response.redirect "index_face.asp"
			response.end
		end if
		rs.close
		set rs=nothing
	end if
	reviewid=checkstr(request.form("reviewid"))
	guestreply=checkstr(request.form("guestreply"))
	if guestreply="" then
		response.write "<script>alert('�ظ�����Ϊ��');history.go(-1);</Script>"
		response.end
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_Review_Table &" where reviewid="&reviewid
	rs.open sql,conn,1,3
	rs("reply")=guestreply
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.redirect "guestbook.asp"
else
	response.write "<script>alert('����Ȩ�ظ������ԣ�');history.go(-1);</Script>"
	response.end
end if
%>