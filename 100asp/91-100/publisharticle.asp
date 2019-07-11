<%@language=vbscript%>
<!--#include virtual="adovbs.inc"-->
<html>
<head>
</head>
<body bgcolor=#c1f7d8>
<center>
<%
dim strArticletitle,strarticlecontent,strarticleauthor,strarticleid
dim strtable,strdsn
if session("name")="" then
    response.write "请你首先登录，才能发表高见"
    response.end
end if
strarticletitle=request.form("title")
strarticlecontent=request.form("content")
strarticleauthor=session("name")
strarticleid=Request.Form("articleid")
strtable="article"
strdsn="dsn=bbs;uid=feng;pwd=feng"
if trim(strarticletitle)="" then
    response.write "主题不能为空"
    response.end
end if
if trim(strarticlecontent)="" then
    strarticletitle=strarticletitle & "(无内容)"
end if
set rs=server.createobject("adodb.recordset")
rs.open strtable,strdsn,3,2
rs.addnew
if request.form("submit")="发表文章" then
    rs("articletitle")=strarticletitle
    rs("articleauthor")=strarticleauthor
    rs("articlecontent")=strarticlecontent
    response.write "文章发表成功"
elseif request.form("submit")="跟贴文章" then
    rs("articletitle")=strarticletitle
    rs("articleauthor")=strarticleauthor
    rs("articlecontent")=strarticlecontent
    rs("articleparent")=strarticleid
end if
rs.update
rs.close
set rs=nothing
%>
<%
'修改跟贴文章数
if request.form("submit")="跟贴文章" then
    strchangesql="update article set articlefellownumber=articlefellownumber+1 where articleid=" & strarticleid
    strconn="dsn=bbs;uid=feng;pwd=feng"
    set conn=server.createobject("adodb.connection")
    conn.open strconn
    set rs=conn.execute(strchangesql)
    set rs=nothing
    conn.close
    set conn=nothing
    response.write "文章成功跟贴"
end if
%>
</body>
</html>

