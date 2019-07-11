
<%
dim founderr
dim errmsg
founderr=false
if trim(request("name"))="" then
	founderr=true
	response.write "<script language=JavaScript>" & chr(13) & "alert('请输入您的昵称!');" & "history.back()" & "</script>" 
elseif trim(request("password"))="" then
	founderr=true
	response.write "<script language=JavaScript>" & chr(13) & "alert('请输入您的密码!');" & "history.back()" & "</script>" 
elseif trim(request("email"))="" then
	founderr=true
	response.write "<script language=JavaScript>" & chr(13) & "alert('请输入您的E_mail!');" & "history.back()" & "</script>" 
elseif request("t_password")<>request("password") then
	founderr=true
	response.write "<script language=JavaScript>" & chr(13) & "alert('两次输入的密码不符!');" & "history.back()" & "</script>" 
end if
if not founderr then
	dim name,pwd,email,person
	name=request("name")
	pwd=request("password")
	email=request("email")
	person=request("person")
	dim current_userRS
	set rs=server.createobject("adodb.recordset")
	set rs1=server.createobject("adodb.recordset")
	conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
	sqladdnew="select * from pass where name='"& name &"'"
	rs.open sqladdnew,conn,1,1
		if RS.eof or RS.bof then
		sqladdnew="insert into pass (name,pwd,email,person) values('"&name&"','"&pwd&"','"&email&"','"&person&"')"
		rs1.open sqladdnew,conn,1,1
		if err.number <> 0 then
			response.write "<script language=JavaScript>" & chr(13) & "alert('对不起,操作出错!');" & "history.back()" & "</script>" 
		else
			response.write  "<script language=JavaScript>" & chr(13) & "alert('用户注册成功!!');" & "history.back()" & "</script>" 
		end if
		else
		response.write "<script language=JavaScript>" & chr(13) & "alert('对不起，此用户已被使用');" & "history.back()" & "</script>" 
	end if
set conn=nothing
end if
%>