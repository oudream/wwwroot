<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp" -->
<!--#include file="chkuser.asp" -->
<!--#include file="char.inc"-->
<!--#include file="ChkURL.asp"-->
<%
NewsID=Request.QueryString("NewsID")
typeid=request.form("typeid")
BigClassid=trim(Request.Form("BigClassid"))
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from "& db_News_Table &" where newsid="&newsid
rs.open sql,conn,3,3
dim title
title=rs("title")
rs.close
set rs=nothing
if BigClassid="" then
%>
<script language=javascript>
history.back()
alert("请选择该文章所属的大类，如果系统中还没有文章大类的话，请在《大类管理》中添加新类别！" )
</script>
<%
Response.End
end if

SmallClassid=trim(Request.Form("SmallClassid"))

set rs13=server.createobject("adodb.recordset")
sql13="select * from "& db_News_Table &" where NewsID="&NewsID
rs13.open sql13,conn,3,3
rs13("bigclassid")=bigclassid
if smallclassid="" then
rs13("smallclassid")=null
else
rs13("smallclassid")=smallclassid
end if
rs13("typeid")=typeid
rs13.update
rs13.close
set rs13=nothing


Response.Redirect "newsedit.asp?NewsID="&NewsID
%>