<%@ Language=VBScript %>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="char.inc"-->
<%
reviewID=checkstr(Request.Form("reviewID"))
NewsID=checkstr(Request.Form("NewsID"))
title=checkstr(Request.Form("title"))
url=trim(checkstr(Request.Form("url")))
Author=trim(checkstr((Request.Form("Author"))))

if author="" then
	response.write "<script>alert('����������������');history.back()</script>"
	Response.End
end if
author=htmlencode(author)

email=trim(Request.Form("email"))
if email="" then
	response.write "<script>alert('����������EMAIL��');history.back()</script>"
	Response.End
end if

if  IsValidEmail(email)=false  then
	response.write "<script>alert('��������ȷ��EMAIL��');history.back()</script>"
	Response.End
end if

if Instr(request("content"),"'")>0 or Instr(request("content"),"script")>0 or Instr(request("content"),"onClick")>0  or Instr(request("content"),"onload")>0 then
	Show_Err("�Բ�����������������ݰ����зǷ��ַ���<br><br><a href='javascript:history.back()'>����</a>")
	Response.End 
end if



content=trim(htmlencode1((request.form("content"))))
content=replace(content,"<p> ","")
content=replace(content,"<P> ","")

dim byte1
byte1=split(byteType,"|")

for i=0 to ubound(byte1)
	content=replace(content,trim(byte1(i)),"***")
next
Response.cookies(Forcast_SN)("content")=content

if content="" then
	response.write "<script>alert('�������������ݣ�');history.back()</script>"
	Response.End
end if

set rs=server.createobject("adodb.recordset")
sql="select * from "& db_News_Table &" where NewsId="&NewsId
rs.open sql,conn,1,3
if rs.eof and rs.bof then
	rs.close
	set rs=nothing
	response.write "<script>alert('�޷��Բ����ڵ����½������ۣ�\n ȷ���Ƿ�Ϊ�Ƿ����ύ��');history.back()</script>"
	response.end
else
	checked=rs("checkked")
	if checked<>1 then
		rs.close
		set rs=nothing
		response.write "<script>alert('����δͨ����ˣ����ܽ������ۣ�');history.back()</script>"
		response.end
	else
		rs("titlesize")=1
		rs.update
		rs.close
		reviewip=Request.ServerVariables("REMOTE_ADDR")
		passed=checkstr(Request.Form("passed"))
	
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_Review_Table &"" 
		rs.open sql,conn,1,3
		rs.addnew
		rs("author")=author
		rs("content")=content
		rs("title")=title
		rs("NewsID")=NewsID
		rs("passed")=passed
		rs("reviewip")=reviewip
		rs("email")=email
		rs("updatetime")=now()
		rs.update
		rs.close
		reviewid=reviewID+1
		set rs=nothing
		Response.cookies(Forcast_SN)("content")=""
	end if
end if
Response.Redirect url
%>