<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
<%on error resume next%>
<%
area=request("area")
sql="select introf from area order by id" 
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
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<!--#include file="top.asp"-->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
    <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td colspan="5"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="1" background="images/2x2.gif"><img src="images/1x2-black.gif" width="1" height="2"></td>
                <td><img src="images/1-x.gif" width="754" height="56"></td>
                <td width="1" background="images/1x2-black.gif"><img src="images/1x2-black.gif" width="1" height="2"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td width="229" valign="top" background="images/left-228x5.gif"> <table width="229" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="229" height="53" background="images/left-1b.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="10%" height="18" valign="bottom">&nbsp; </td>
                      <td width="8%" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>客 
                        户 登 录</b></font></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td valign="top"> <table width="228" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td background="images/left-1-left-2b.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td width="21" valign="top"><img src="images/left-1-left-1b.gif" width="21" height="190"></td>
                            <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td> 
                                    <!--#include file="userlogin.asp"-->
                                  </td>
                                </tr>
                                <tr> 
                                  <td> 
                                    <!--#include file="userhelp.asp" -->
                                  </td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td><img src="images/left-1-left-3b.gif" width="229" height="12"></td>
                    </tr>
                  </table></td>
              </tr>
            </table> 
            <table width="229" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="229" height="28" background="images/left-2.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="10%" height="18" valign="bottom">&nbsp; </td>
                      <td width="8%" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>客 
                        户 中 心</b></font></td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <TABLE width=100% border=0 cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
              <TBODY>
                <TR> 
                  <TD height="2"
				vAlign=top 
                style="BORDER-left: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                </TR>
              </TBODY>
            </TABLE
            >
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td>
                  <!--#include file="support.asp" -->
                </td>
              </tr>
            </table></td>
          <td width="7">&nbsp;</td>
          <td width="513" valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td rowspan="2" width="1"></td>
                <td> <table width="504" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td bgcolor="#FFFFFF"> <TABLE cellSpacing=0 cellPadding=0 width=100% border=0>
                          <TBODY>
                            <TR> 
                              <TD width="34%" vAlign=center style="BORDER-left: #CCCCCC 1px solid; BORDER-RIGHT: #CCCCCC 1px solid; BORDER-BOTTOM: #CCCCCC 1px solid;BORDER-TOP: #CCCCCC 1px solid"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td background="images/host-2.gif"><img src="images/allprd.gif" width="153" height="25"></td>
                                  </tr>
                                </table>
                                <br>
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td> <table width="488" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                                        <tr> 
                                          <td width="10%" height="40" align="center" valign="middle" bgcolor="#f5f5f5">编号</td>
                                          <td width="40%" height="40" align="center" bgcolor="#f5f5f5">品&nbsp;&nbsp;名</td>
                                          <td width="20%" align="center" bgcolor="#f5f5f5">类型</td>
                                          <td width="20%" height="40" align="center" bgcolor="#f5f5f5"> 
                                            价&nbsp;&nbsp;格</td>
                                          <td width="10%" align="center" bgcolor="#f5f5f5">详情</td>
                                        </tr>
                                        <%
sql="select * from subs order by id" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
n=0
'预留分页显示功能
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

amount=rs.recordcount
rs.pagesize=30
rs.AbsolutePage=pagecount

do while not rs.eof
%>
                                        <tr> 
                                          <td height="40" align="center" valign="middle" bgcolor="#FFFFFF"><a href="showsubs.asp?id=<%=rs("id")%>" target="_blank"><%=rs("id")%></a></td>
                                          <td height="40" bgcolor="#FFFFFF"><a href="showsubs.asp?id=<%=rs("id")%>" target="_blank"><%=rs("subsname")%></a></td>
                                          <td height="40" bgcolor="#FFFFFF"><%=rs("area")%></td>
                                          <td bgcolor="#FFFFFF"> <%if rs("ifdisc")=true then%> <%response.write FormatNumber(csng(rs("price"))*session("y_it_userdiscount"),2)%> <%else%> <%response.write FormatNumber(csng(rs("price")),2)%> <%end if%>
                                            元/<%=rs("subsunit")%></td>
                                          <td align="center" bgcolor="#FFFFFF"><a href="showsubs.asp?id=<%=rs("id")%>" target="_blank">详情</a></td>
                                        </tr>
                                        <!--   要不要            <tr> 
               <td height="6" colspan="2"><div align="center"><a href="#" onClick="window.open('add.asp?add=<%=rs("bookbm")%>','add','scrollbars=yes,resizable=no,width=650,height=450 top=10,left=10')"><img src="image/in1a.gif" width="170" height="24" border="0"></a><a href="showsubs.asp?id=<%=rs("id")%>" target="_blank"><img src="image/in1b.gif" width="107" height="24" border="0"></a></div></td> 
              </tr>     -->
                                        <%
rs.movenext
n=n+1                                                                     
if n>=rs.pagesize then exit do                                                           
loop
%>
                                      </table></td>
                                  </tr>
                                  <tr> 
                                    <td>&nbsp; <div align="right">
                                        <table width="488" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#f5f5f5">
                                          <tr> 
                                            <td width="40%" height="20">
											
<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr><td></td><td>            
<%
zj=1
for zi=1 to rs.pagecount
%> 
   <a href="showallproduct.asp?page=<%=zj%>">|&nbsp;<%=zj%>&nbsp;|</a>         
<%
if zj mod 12 =0 then response.Write("<br><br>")
zj=zj+1
next
%> 
</td></tr>
</table> 
									
											
											
											</td>
                                            <td width="60%" height="20" align="right">
											
											
											
											
                                        共&nbsp;<%=amount%>&nbsp;个&nbsp;&nbsp;每页&nbsp;30&nbsp;个&nbsp;&nbsp;<%=pagecount%></font>/<%=rs.pagecount%> 
                                        <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%>
                                        <a href="showallproduct.asp?page=<%=cstr(pagecount+1)%>"><font color="#ff0000">下一页</font></a><font color="#ff0000"> 
                                        <% end if %>
                                        <% if rs.pagecount>1 and rs.pagecount=pagecount then %>
                                        <a href="showallproduct.asp?page=<%=cstr(pagecount-1)%>"> 
                                        上一页</a> 
                                        <%end if%>
                                        <% if pagecount<>1 and rs.pagecount<>pagecount then%>
                                        <a href="showallproduct.asp?page=<%=cstr(pagecount-1)%>">上一页</a> 
                                        <a href="showallproduct.asp?page=<%=cstr(pagecount+1)%>"><font color="#ff0000">下一页</font></a><font color="#ff0000"> 
                                        <% end if%>											</td>
                                          </tr>
                                        </table>
<%
rs.close
set rs=nothing
%>
                                     </td>
                                  </tr>
                                </table>
                                <br>
                              </TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <br> </td>
                    </tr>
                  </table>
                  
                </td>
              </tr>
            </table></td>
          <td width="7" background="images/right-7x5.gif"><img src="images/right-7x5.gif" width="7" height="5"></td>
        </tr>
      </table></td>
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
  </tr>
</table>
<!--#include file="copyright.asp"-->
</body>
</html>
