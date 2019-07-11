<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%on error resume next%>
<%
gys=request("gys")
sql="select gysxq from gys where gys='"&gys&"' order by id" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
gysxq=rs("gysxq")
rs.close
set rs=nothing
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-<%=gys%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/southdns.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<!--#include file="top.asp"--> 
<table width="750" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width=192 valign="top">
<!--#include file="userlogin.asp"-->
<!--#include file="dynamic.asp"-->
<!--#include file="cserv.asp"-->
<!--#include file="link.asp"-->
      </td>
    <td width="30">&nbsp;</td>
    <td valign="top" width="576"> 
      <div align="right">
        <table width="563" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="0"></td>
          </tr>
          <tr> 
            <td height="25" bgcolor="#9CCCFF" style="BORDER-RIGHT: #0079ce 1px solid; BORDER-TOP: #0079ce 1px solid; BORDER-LEFT: #0079ce 1px solid; BORDER-BOTTOM: #0079ce 1px solid"><p
                         align="center" style="margin-top:3px;">
              <div align="center">=== -<%=gys%> ===</div>
            </td>
          </tr>
          <tr> 
            <td><br>&nbsp;&nbsp;&nbsp;&nbsp;<%  
gysxq=replace(gysxq,chr(13),"<br>")
gysxq=replace(gysxq,chr(32),"&nbsp;")
response.write gysxq
%><br><br></td>
          </tr>
        </table>

<%
sql="select * from subs where gys='"&gys&"' order by id" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
n=0
'预留分页显示功能
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=6
rs.AbsolutePage=pagecount

do while not rs.eof
%>
<table border="0" cellpadding="0" cellspacing="0">
    <tr>
        <td width="0" height="1" colspan="3" bgcolor="#9CCFFF"></td>
    </tr>
    <tr>
        <td width="1" bgcolor="#9CCFFF"></td>
        <td width="280" valign="top"><table border="0" cellpadding="0"
             cellspacing="0">
                <tr>
                    <td width="277" height="88" valign="top"><table border="0" cellpadding="0" cellspacing="0">
    <tr>
        <td width="120" height="120" rowspan="4"><div align="center"><img src=showimage.asp?id=<%=rs("id")%>&tu=stu border="0" width="68" height="98"></div></td>
        <td width="157" height="30">品&nbsp;&nbsp;名：<%=rs("subsname")%></td>
    </tr>
    <tr>
        <td width="157" height="30">价&nbsp;&nbsp;格：
<%if rs("ifdisc")=true then%>
<%response.write FormatNumber(csng(rs("price"))*session("y_it_userdiscount"),2)%>
<%else%>
<%response.write FormatNumber(csng(rs("price")),2)%>
<%end if%>
 RMB /<%=rs("subsunit")%></td>
    </tr>
    <tr>
        <td width="157" height="30">服务商：<%=rs("gys")%></td>
    </tr>
    <tr>
        <td width="157" height="60">简&nbsp;&nbsp;介：
<%if len(rs("other"))>25 then%> <%=left(rs("other"),22)%>...<%else%> 
<%=rs("other")%> <%end if%><br><br></td>
    </tr>
    <tr>
        <td width="277" height="6" colspan="2"><a href="#" onClick="window.open('add.asp?add=<%=rs("bookbm")%>','add','scrollbars=yes,resizable=no,width=650,height=450 top=10,left=10')"><img src="image/in1a.gif" width="170" height="24" border="0"></a><a href="showsubs.asp?id=<%=rs("id")%>"><img src="image/in1b.gif" width="107" height="24" border="0"></a></td>
    </tr>
</table></td>
                  
<%
rs.movenext
if rs.eof or rs.bof then 
response.write "</tr></table> </td><td width=1 bgcolor=#9CCFFF></td></tr></table>"
exit do
else
n=n+1
end if
%><td width=7 valign=bottom><img src=image/inbgnew.gif></td>       <td width="277" valign="top"><table border="0" cellpadding="0" cellspacing="0">
    <tr>
        <td width="120" height="120" rowspan="4"><div align="center"><img src=showimage.asp?id=<%=rs("id")%>&tu=stu border="0" width="68" height="98"></div></td>
        <td width="160" height="30">品&nbsp;&nbsp;名：<%=rs("subsname")%></td>
    </tr>
    <tr>
        <td width="157" height="30">价&nbsp;&nbsp;格：
<%if rs("ifdisc")=true then%>
<%response.write FormatNumber(csng(rs("price"))*session("y_it_userdiscount"),2)%>
<%else%>
<%response.write FormatNumber(csng(rs("price")),2)%>
<%end if%>
 RMB /<%=rs("subsunit")%></td>
    </tr>
    <tr>
        <td width="157" height="30">服务商：<%=rs("gys")%></td>
    </tr>
    <tr>
        <td width="157" height="60">简&nbsp;&nbsp;介：
<%if len(rs("other"))>25 then%> <%=left(rs("other"),22)%>...<%else%> 
<%=rs("other")%> <%end if%><br><br></td>
    </tr>
    <tr>
        <td width="277" height="6" colspan="2"><a href="#" onClick="window.open('add.asp?add=<%=rs("bookbm")%>','add','scrollbars=yes,resizable=no,width=650,height=450 top=10,left=10')"><img src="image/in1a.gif" width="170" height="24" border="0"></a><a href="showsubs.asp?id=<%=rs("id")%>"><img src="image/in1b.gif" width="107" height="24" border="0"></a></td>
    </tr>
</table></td>
                </tr>
            </table> </td>
        <td width="1" bgcolor="#9CCFFF"></td>
    </tr>
</table>
<%
rs.movenext
n=n+1                                                                     
if n>=rs.pagesize then exit do                                                           
loop
%>
        <table width="563" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td rowspan="2" width="0" bgcolor="#3D5E7D"></td>
            <td height="149" valign="top"> 
              <table width="500" border="0" cellspacing="0" cellpadding="0" align="center">
                <tr> 
                  <td> 
                    <div align="right"><font color=red><%=pagecount%></font>/<%=rs.pagecount%> 
                      <img src="image/in2_r14_c8.gif" width="36" height="20"> 
                      <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%> 
                      <a href="showgys.asp?gys=<%=gys%>&page=<%=cstr(pagecount+1)%>"><font color="#ff0000">下一页</a> 
                      <% end if %> <% if rs.pagecount>1 and rs.pagecount=pagecount then %> 
                      <a href="showgys.asp?gys=<%=gys%>&page=<%=cstr(pagecount-1)%>"> 
                      上一页</a> <%end if%> <% if pagecount<>1 and rs.pagecount<>pagecount then%> 
                      <a href="showgys.asp?gys=<%=gys%>&page=<%=cstr(pagecount-1)%>">上一页</a> 
                      <a href="showgys.asp?gys=<%=gys%>&page=<%=cstr(pagecount+1)%>"><font color="#ff0000">下一页</a>  
                      <% end if%>&nbsp;</div>
                  </td>
                </tr>
<%
rs.close
set rs=nothing
%> 
<%
sql="select * from message where newstype='bill'" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
%> 
                <tr>
                  <td height="25"><b><%=rs("subject")%></b></td>
                </tr>
                
                <tr> 
                  <td><%=rs("message")%>
               </td>
                </tr>
<%
rs.close
set rs=nothing
%>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="1"></td>
          </tr>
        </table><br>
      </div>
    </td>
  </tr>
</table>
<!--#include file="copyright.asp"-->
</body>
</html>
