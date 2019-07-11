<%
dim IID,cate
if session("IID")="" then
IID=1
cate="办公室"
else
	IID=session("IID")
	cate=session("cate")
end if
dim name,sex,borndate,term,func
if request("text1")="" then
	response.write "<script language=JavaScript>" & chr(13) & "alert('请输入用户名！');" & "history.back()" & "</script>" 
end if

if request("text2")=""  then
	response.write "<script language=JavaScript>" & chr(13) & "alert('出生年月不能为空！');" & "history.back()" & "</script>" 
end if
borndate=request("text2")
if isdate(borndate)=0 then
	response.write "<script language=JavaScript>" & chr(13) & "alert('出生年月输入错误，将返回前页！');" & "history.back()" & "</script>" 
end if 

if request("text3")=""  then
	response.write "<script language=JavaScript>" & chr(13) & "alert('职务不能为空！');" & "history.back()" & "</script>" 
end if
name=request("text1")
func=request("text3")
term=request("text4")
sex=request("sel1")
dim sql
dim rsadd
dim connstr
set rs=server.createobject("adodb.recordset") 
conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
sql="insert into mytable(name,borndate,sex,term,func,cateid,cate) values('"&name&"','"&borndate&"','"&sex&"','"&term&"','"&func&"','"&IID&"','"&cate&"')"
rs.Open sql,conn,1,1
Response.Write "记录添加成功!"
%>