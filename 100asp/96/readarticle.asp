<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY bgcolor=#c1f7d8>
<p align=center><font size=5>�Ķ�����</font></p>
<%
'��ʾ��������
dim strDsn,strSelectSql
strSelectSql="select * from article where articleid=" & Request.QueryString("id")
strDsn="Dsn=bbs;uid=feng;pwd=feng"
set rs=server.CreateObject ("adodb.recordset")
rs.Open strselectsql,strdsn,3,1
%>
<table border=1 width=100%>
  <tr>
    <td align=center width=25%>ʱ��</td>
    <td align=center width=49%>����</td>
    <td align=center width=10%>����</td>
    <td align=center width=8%>����</td>
    <td align=center width=8%>����</td>
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
��������:
<br><br>
<%=rs("articlecontent")%>
<%
rs.close
set rs=nothing
%>
<%
'�޸ı�������
strdsn="dsn=bbs;uid=feng;pwd=feng"
strchangesql="update article set articleaccessnumber=articleaccessnumber+1 where articleid=" & Request.QueryString("id")
set changers=server.createobject("adodb.recordset")
changers.open strChangesql,strdsn,1,3
set changers=nothing
%>
<%
'��ʾ��������
strselectsql="select * from article where articleparent=" & request.querystring("id")
strdsn="dsn=bbs;uid=feng;pwd=feng"
set rs=server.createobject("adodb.recordset")
rs.open strselectsql,strdsn,3,1
rs.pagesize=10
nextpage=request.form("nextpage")
if nextpage="" then
    session("abspage")=1
else
    if nextpage="��һҳ" then
        session("abspage")=session("abspage")-1
    elseif nextpage="��һҳ" then
        session("abspage")=session("abspage")+1
    elseif nextpage="��һҳ" then
        session("abspage")=1
    elseif nextpage="���һҳ" then
        session("abspage")=rs.pagecount
    end if
    rs.absolutepage=session("abspage")
end if    
    if rs.recordcount>0 then
        i=0
        response.write "<table border=1 width=100%/>"
        response.write "<tr><td colspan=5 align=center>"
        response.write "����" & rs.Recordcount & "������"
%>
     <tr>
       <td align=center width=25%>ʱ��</td>
       <td align=center width=49%>����</td>
       <td align=center width=10%>����</td>
       <td align=center width=8%>����</td>
       <td align=center width=8%>����</td>
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
      response.write "<input type=submit value=��һҳ name=nextpage>"
   end if
   if (session("abspage"))<rs.pagecount then
      response.write "<input type=submit value=��һҳ name=nextpage>"
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
  <td>����:</td>
  <td><input type=tex name=title size=61></td>
  </tr>
  <tr>
  <cd colspan=2>����:</td>
  </tr>
  <tr>
  <td colspan=2><textarea id=textarea1 name=content style="Height:100px;width:500px"></textarea></td>
  </tr>
</table>
<center><br>
<input id=submit1 name=submit type=submit value=��������>
<input id=submit1 name=submit type=submit value=��������>
<input id=reset1 name=reset1 type=reset value=��д����>
</center>
</form>
</BODY>
</HTML>
