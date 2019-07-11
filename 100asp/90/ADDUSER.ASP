<%
dim name,pwd,email,person
dim sql
dim rs
if request("txt1")="" then
	response.write "<script language=JavaScript>" & chr(13) & "alert('请输入用户名！');" & "history.back()" & "</script>" 
end if
if request("txt2")=""  then
	response.write "<script language=JavaScript>" & chr(13) & "alert('密码不充许为空！');" & "history.back()" & "</script>" 
end if
if request("txt2")<>request("txt3") then
	response.write "<script language=JavaScript>" & chr(13) & "alert('两次输入的密码不一致！');" & "history.back()" & "</script>" 
end if
name=request("txt1")
password=request("txt2")
email=request("txt4")
sex=request("sel1")
set rs=server.createobject("adodb.recordset")
set rs1=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="select * from user where username='"&name&"'"
rs.Open sql,conn,1,1
if rs.EOF then
	sql="insert into user(username,password,email,sex) values('"&name&"','"&password&"','"&email&"','"&sex&"')"
	rs1.Open sql,conn,1,1
	response.write "<script language=JavaScript>" & chr(13) & "alert('用户注册成功！');"&"window.location.href = 'index.asp'"&" </script>" 
else
	response.write "<script language=JavaScript>" & chr(13) & "alert('此用户已有人使用！');" & "history.back()" & "</script>" 
end if
%>
