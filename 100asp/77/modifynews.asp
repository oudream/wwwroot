<%
dim IID
IID=request("id")
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
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="update mytable set username='"&username&"',sex='"&sex&"',borndate='"&borndate&"',phone='"&phone&"',BP='"&BP&"',policy='"&policy&"',address='"&addess&"',company='"&company&"',email='"&email&"' where id="+IID
rs.Open sql,conn,1,1
response.write "<script language=JavaScript>" & chr(13) & "alert('同学信息修改成功！');"&"window.location.href = 'index.asp'"&" </script>" 
set rs=nothing
%>
