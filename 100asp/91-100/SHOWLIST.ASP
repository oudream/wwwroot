<html>
<head>
<title>浏览标题</title>
</head>
<body bgcolor=#c1f7d8>
<p align=center><font size=5>浏览标题</font></p>
<%
dim strdsn,strselectsql
dim intarticleid
if request.querystring("ID")="" then
    intArticleid=0
end if
strselectsql="select * from article where articleparent=" & intarticleid
strdsn="dsn=bbs;uid=feng;pwd=feng"
set rs=server.createobject("adodb.recordset")
rs.open strselectsql,strdsn,3,1
rs.pagesize=10
nextpage=request.form("nextpage")
if nextpage="" then
    session("abspage")=1
else
    if nextpage="上一页" then
        session("abspage")=session("abspage")-1
    elseif nextpage="下一页" then
        session("abspage")=session("abspage")+1
    elseif nextpage="第一页" then
        session("abspage")=1
    elseif nextpage="最后一页" then
        session("abspage")=rs.pagecount
    end if
    rs.absolutepage=session("abspage")
end if

if rs.recordcount>0 then
    i=0
    response.write "<table border=1 width=100%/>"
    response.write "<tr><td colspan=5 align=center>"
    if intarticleid=0 then
        response.write "这是第" & session("abspage") & "页" & "&nbsp;共有" & rs.recordcount & "个主题"
    else
        responsew.write "共有" & rs.recordcount & "个跟贴"
    end if
    response.write "</tr>"
%>
<tr>
<td align=center width=25%>时间</td>
<td align=center width=49%>主题</td>
<td align=center width=10%>作者</td>
<td align=center width=8%>被读</td>
<td align=center width=8%>跟贴</td>
</tr>
<%
do while not rs.eof and i<10
%>
<tr>
<td align=center><%=rs("articledate")%>&nbsp;<%=rs("articletime")%></td>
<td align=center><a href=readarticle.asp?id=<%=rs("articleid")%>><%=rs("articletitle")%></a></td>
<%if session("name")="" then%>
    <td align=center><%=rs("articleauthor")%></td>
<%else%>
    <td align=center><a href=newmessage.asp?toname=<%=rs("articleauthor")%>><%=rs("articleauthor")%></td>
<%end if%>
<td align=center><%=rs("articleaccessnumber")%></td>
<td align=center><%=rs("articlefellownumber")%></td>
</tr>
<%rs.movenext
i=i+1
loop
response.write "</table></center>"
response.write "<center><form action showlist.asp method=post>"
if rs.pagecount>1 then
    response.write "<input type=submit value=上一页 name=nextpage>"
end if
if(session("abspage"))<rs.pagecount then
    response.write "<input type=submit value=下一页 name=nextpage>"
end if
'end if
response.write "</form>"
end if
rs.close
set rs=nothing
%>
</tr>
</table>
</body>
</html>
