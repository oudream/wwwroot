<%@ Language=VBScript %>
<%
dim rs
dim founderr
if trim(request("textname")) = "" then
	founderr = true
  response.write "<script language=JavaScript>" & chr(13) & "alert('请输入昵称！');" & "history.back()" & "</script>" 
end if
if trim(request("textpassword")) = "" then
	founderr = true
  response.write "<script language=JavaScript>" & chr(13) & "alert('请输入密码！');" & "history.back()" & "</script>" 
end if
if trim(request("textcontent")) = "" then
	founderr = true
  response.write "<script language=JavaScript>" & chr(13) & "alert('请输入内容！');" & "history.back()" & "</script>" 
end if
if founderr then
	response.write errmsg
else
	dim name,pwd,content,publish
	publishtime=now()
	name=Request("textname")
	pwd=request("textpassword")
	content=Request("textContent")
	dim checksql
	set rs=server.CreateObject("ADODB.RecordSet")
	set conn=server.createobject("ADODB.CONNECTION")
	connstr = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
	conn.Open connstr
	checksql="select * from pass where name='"& name &"' and pwd='"& pwd &"'"
	set rs=conn.execute(checksql)
	if rs.eof or rs.bof then
	response.write "<script language=JavaScript>" & chr(13) & "alert('您还没有注册或名称和密码有错误，请先验证！！');" & "history.back()" & "</script>" 
 	else
	checksql="insert into message(user,content,publishtime) values('"&name&"','"&content&"','"&publishtime&"')"
	set rs=conn.Execute(checksql)
	if err.number <> 0 then
	response.write "<script language=JavaScript>" & chr(13) & "alert('无法保存,数据库操作出错:！');" & "history.back()" & "</script>" 
	else	
		response.write "<script language=JavaScript>" & chr(13) & "alert('提交成功！');" & "history.back()" & "</script>" 
	end if
	end if
set rs=nothing 
set conn=nothing
end if
%>