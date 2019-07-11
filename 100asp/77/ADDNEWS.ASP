<%
dim name,pwd,email,person
dim sql
dim rs
if request("text1")="" then
	response.write "<script language=JavaScript>" & chr(13) & "alert('请输入用户名！');" & "history.back()" & "</script>" 
	Response.End
end if
borndate=request("text2")
if request("text2")<>"" then
	if isdate(borndate)=0 then 
	response.write "<script language=JavaScript>" & chr(13) & "alert('生日输入错误，将返回前页！');" & "history.back()" & "</script>" 
	response.end
	end if
end if
username=request("text1")
sex=request("sel1")
phone=request("text3")
BP=request("text4")
policy=request("text5")
address=request("text6")
company=request("text7")
email=request("text8")
set rs=server.createobject("adodb.recordset")
set rs1=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="select * from mytable where username='"&username&"'"
rs.Open sql,conn,1,1
if rs.EOF or rs.bof then
	sql="insert into mytable(username,sex,borndate,phone,BP,policy,address,company,email) values('"&username&"','"&sex&"','"&borndate&"','"&phone&"','"&BP&"','"&policy&"','"&address&"','"&company&"','"&email&"')"
	rs1.Open sql,conn,1,1
	response.write "<script language=JavaScript>" & chr(13) & "alert('同学信息添加成功！');"&"window.location.href = 'index.asp'"&" </script>" 
	set rs1=nothing
else
	response.write "<script language=JavaScript>" & chr(13) & "alert('此同学早已添加！');" & "history.back()" & "</script>" 
end if
rs.Close
set rs=nothing
%>
