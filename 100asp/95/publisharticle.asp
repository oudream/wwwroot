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
    response.write "�������ȵ�¼�����ܷ���߼�"
    response.end
end if
strarticletitle=request.form("title")
strarticlecontent=request.form("content")
strarticleauthor=session("name")
strarticleid=Request.Form("articleid")
strtable="article"
strdsn="dsn=bbs;uid=feng;pwd=feng"
if trim(strarticletitle)="" then
    response.write "���ⲻ��Ϊ��"
    response.end
end if
if trim(strarticlecontent)="" then
    strarticletitle=strarticletitle & "(������)"
end if
set rs=server.createobject("adodb.recordset")
rs.open strtable,strdsn,3,2
rs.addnew
if request.form("submit")="��������" then
    rs("articletitle")=strarticletitle
    rs("articleauthor")=strarticleauthor
    rs("articlecontent")=strarticlecontent
    response.write "���·���ɹ�"
elseif request.form("submit")="��������" then
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
'�޸ĸ���������
if request.form("submit")="��������" then
    strchangesql="update article set articlefellownumber=articlefellownumber+1 where articleid=" & strarticleid
    strconn="dsn=bbs;uid=feng;pwd=feng"
    set conn=server.createobject("adodb.connection")
    conn.open strconn
    set rs=conn.execute(strchangesql)
    set rs=nothing
    conn.close
    set conn=nothing
    response.write "���³ɹ�����"
end if
%>
</body>
</html>

