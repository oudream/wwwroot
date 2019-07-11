<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%on error resume next%>
<%
area=request("area")
sql="select introf from area where area='"&area&"' order by id" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
introf=rs("introf")
rs.close
set rs=nothing
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-<%=area%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/southdns.css" type="text/css">
</head>
 
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td colspan="2">
      <!--#include file="top.asp"-->
    </td>
  </tr>
  <tr> 
    <td valign="top"> <table width="250" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>&nbsp;<!--#include file="userlogin.asp" -->
</td>
        </tr>
        <tr>
          <td>&nbsp;<!--#include file="dynamic.asp" -->
</td>
        </tr>
        <tr>
          <td>&nbsp;<!--#include file="cserv.asp" -->
</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table>
      <p>&nbsp;</p>
      </td>
    <td> <table width="504" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="504"> <%=area%> 
            <%  
introf=replace(introf,chr(13),"<br>")
introf=replace(introf,chr(32),"&nbsp;")
response.write introf
%>
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <%
sql="select * from subs where area='"&area&"' order by id" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
n=0
'预留分页显示功能
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=10
rs.AbsolutePage=pagecount

do while not rs.eof
%>
        <tr> 
          <td> <table width="500" border="1" cellpadding="1" cellspacing="1">
              <tr> 
                <td width="171" height="120" rowspan="4"><div align="center"><img src=showimage.asp?id=<%=rs("id")%>&tu=stu border="0" width="68" height="98"></div></td>
                <td width="329" height="30">品&nbsp;&nbsp;名：<%=rs("subsname")%></td>
              </tr>
              <tr> 
                <td width="329" height="30">价&nbsp;&nbsp;格： 
                  <%if rs("ifdisc")=true then%>
                  <%response.write FormatNumber(csng(rs("price"))*session("y_it_userdiscount"),2)%>
                  <%else%>
                  <%response.write FormatNumber(csng(rs("price")),2)%>
                  <%end if%>
                  RMB /<%=rs("subsunit")%></td>
              </tr>
              <tr> 
                <td width="329" height="30">服务商：<%=rs("gys")%></td>
              </tr>
              <tr> 
                <td width="329" height="60">简&nbsp;&nbsp;介： 
                  <%if len(rs("other"))>25 then%>
                  <%=left(rs("other"),22)%>... 
                  <%else%>
                  <%=rs("other")%> 
                  <%end if%>
                  <br> <br></td>
              </tr>
              <tr> 
                <td height="6" colspan="2"><div align="center"><a href="#" onClick="window.open('add.asp?add=<%=rs("bookbm")%>','add','scrollbars=yes,resizable=no,width=650,height=450 top=10,left=10')"><img src="image/in1a.gif" width="170" height="24" border="0"></a><a href="showsubs.asp?id=<%=rs("id")%>" target="_blank"><img src="image/in1b.gif" width="107" height="24" border="0"></a></div></td>
              </tr>
            </table></td>
        </tr>
        <%
rs.movenext
n=n+1                                                                     
if n>=rs.pagesize then exit do                                                           
loop
%>
        <tr> 
          <td>&nbsp; <div align="right"><font color=red><%=pagecount%></font>/<%=rs.pagecount%> 
              <img src="image/in2_r14_c8.gif" width="36" height="20"> 
              <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%>
              <a href="showarea.asp?area=<%=area%>&page=<%=cstr(pagecount+1)%>"><font color="#ff0000">下一页</font></a><font color="#ff0000"> 
              <% end if %>
              <% if rs.pagecount>1 and rs.pagecount=pagecount then %>
              <a href="showarea.asp?area=<%=area%>&page=<%=cstr(pagecount-1)%>"> 
              上一页</a> 
              <%end if%>
              <% if pagecount<>1 and rs.pagecount<>pagecount then%>
              <a href="showarea.asp?area=<%=area%>&page=<%=cstr(pagecount-1)%>">上一页</a> 
              <a href="showarea.asp?area=<%=area%>&page=<%=cstr(pagecount+1)%>"><font color="#ff0000">下一页</font></a><font color="#ff0000"> 
              <% end if%>
              &nbsp;</font></font></div></td>
        </tr>
      </table>
      <%
rs.close
set rs=nothing
%>
      <%
sql="select * from message where newstype='bill'" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
%>
      <b><%=rs("subject")%></b><%=rs("message")%> 
      <%
rs.close
set rs=nothing
%>
    </td>
  </tr>
  <tr> 
    <td colspan="2">
      <!--#include file="copyright.asp"-->
    </td>
  </tr>
</table>
</body>
</html>
