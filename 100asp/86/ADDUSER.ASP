<%@ Language=VBScript %>
<%
dim name,pwd,email,person
dim sql
dim rs1,rs
name=request("txt1")
password=request("txt2")
address=request("txt3")
phone=request("txt4")&"-"&request("txt5")
email=request("txt6")
age=request("sel1")
set rs=server.createobject("adodb.recordset")
set rs1=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="select * from user where username='" & name & "'"
rs.Open sql,conn,1,1
if rs.EOF then
	sql="insert into user(username,age,address,email,password,phone) values('"&name&"','"&age&"','"&address&"','"&email&"','"&password&"','"&phone&"')"
	rs1.Open sql,conn,1,1
	session("name")=name
	Response.Write "用户注册成功!"
	Response.Redirect "index.asp"
else
	Response.Write "此用户已有人使用!"
end if
%>