<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY bgcolor=#c1f7d8>
<p align=center><font size=5>阅读文章</font></p>
<%
'显示文章内容
dim strDsn,strSelectSql
strSelectSql="select * from article where articleid=" & Request.QueryString("id")
strDsn="Dsn=bbs;uid=feng;pwd=feng"
set rs=server.CreateObject ("adodb.recordset")
rs.Open strselectsql,strdsn,3,1
%>
<table border=1 width=100%>
  <tr>
    <td align=center width=25%>时间</td>
    <td align=center width=49%>主题</td>
    <td align=center width=10%>作者</td>
    <td align=center width=8%>被读</td>
    <td align=center width=8%>跟贴</td>
  </tr>
  <tr>
    <td align=center><%=rs("articledate")%>&nbsp;<%=rs("articletime")%></td>
    <td align=center><%=rs("articletitle")%>
    <td align=center><%=rs("articleauthor")%>
    <td align=center><%=rs("articleaccessnumber")%>
    <td align=center><%=rs("articlefellownumber")%>
  </tr>
</table>
<br>
文章内容:
<br><br>
<%=rs("articlecontent")%>
<%
rs.close
set rs=nothing
%>
<%
'修改被读次数
strdsn="dsn=bbs;uid=feng;pwd=feng"
strchangesql="update article set articleaccessnumber=articleaccessnumber+1 where articleid=" & Request.QueryString("id")
set changers=server.createobject("adodb.recordset")
changers.open strChangesql,strdsn,1,3
set changers=nothing
%>
<%
'显示跟贴文章
strselectsql="select * from article where articleparent=" & request.querystring("id")
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
        response.write "共有" & rs.Recordcount & "个跟贴"
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
    <td align=center><%=rs("articletitle")%>
    <td align=center><%=rs("articleauthor")%>
    <td align=center><%=rs("articleaccessnumber")%>
    <td align=center><%=rs("articlefellownumber")%>
  </tr>
<% rs.movenext
   i=i+1
   loop
   response.write "</table></center>"
   response.write "<center><form action showlist.asp method=post>"
   if rs.pagecount>1 then
   if (session("abspage"))>1 then
      response.write "<input type=submit value=上一页 name=nextpage>"
   end if
   if (session("abspage"))<rs.pagecount then
      response.write "<input type=submit value=下一页 name=nextpage>"
   end if
   end if
   response.write "</form>"
end if
rs.close
set rs=nothing
%>
</tr>
</table>
<hr>
<form action=publisharticle.asp method=post>
<input type=hidden name=articleid value=<%=Request.QueryString("id")%>>
<table border=0>
  <tr>
  <td>主题:</td>
  <td><input type=tex name=title size=61></td>
  </tr>
  <tr>
  <cd colspan=2>内容:</td>
  </tr>
  <tr>
  <td colspan=2><textarea id=textarea1 name=content style="Height:100px;width:500px"></textarea></td>
  </tr>
</table>
<center><br>
<input id=submit1 name=submit type=submit value=跟贴文章>
<input id=submit1 name=submit type=submit value=发表文章>
<input id=reset1 name=reset1 type=reset value=重写文章>
</center>
</form>
</BODY>
</HTML>
