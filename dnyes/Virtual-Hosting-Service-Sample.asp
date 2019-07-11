<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
<%on error resume next%>
<%
area="虚拟主机"
%>
<html>
<head>
<title>信网 数信科技-虚拟主机 建站空间 具有良好的速度、稳定性、可靠性、安全性，是企业信息发布和实现电子商务的网站平台</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="www.dnyes.com-host,虚拟主机,建站空间,具有良好的速度,稳定性,可靠性,安全性,是企业信息发布和实现电子商务的网站平台">
<meta name="description" content="www.dnyes.com-host,虚拟主机,建站空间,具有良好的速度,稳定性,可靠性,安全性,是企业信息发布和实现电子商务的网站平台">
</head>
 
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<!--#include file="top.asp"-->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
    <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
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
          <td width="229" valign="top" background="images/left-228x5.gif"> <TABLE width=100% border=0 cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
              <TBODY>
                <TR> 
                  <TD height="10" 
                vAlign=top 
                style="BORDER-left: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                </TR>
              </TBODY>
            </TABLE>
            <table width="229" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="229" height="28" background="images/left-2.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="10%" height="18" valign="bottom">&nbsp; </td>
                      <td width="8%" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>虚拟主机客户展示</b></font></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td> <TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
                    <TBODY>
                      <TR> 
                        <TD height="2" 
                vAlign=top 
                style="BORDER-left: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td><!--#include file="aboutarea.asp"--></td>
                    </tr>
                  </table>
                  <TABLE width=100% border=0 cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
                    <TBODY>
                      <TR> 
                        <TD height="1"
				vAlign=top 
                style="BORDER-left: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                      </TR>
                    </TBODY>
                  </TABLE></td>
              </tr>
            </table>
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td><table width="229" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="229" height="28" background="images/left-2.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td width="10%" height="18" valign="bottom">&nbsp; </td>
                            <td width="8%" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                            <td width="82%" valign="bottom"><font color="#FFFFFF"><b>主 机 优 势 
                             </b></font></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
			  <tr>	<td align="center"><!--#include file="dnyesserver.asp" -->
		</td></tr>
            </table>
            
          </td>
          <td width="7">&nbsp;</td>
          <td width="513" valign="top"> 
            <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td rowspan="2" width="1"></td>
                <td>
<table width="504" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td bgcolor="#FFFFFF"> <TABLE cellSpacing=0 cellPadding=0 width=100% border=0>
                          <TBODY>
                            <TR> 
                              <TD width="34%" vAlign=center style="BORDER-left: #CCCCCC 1px solid; BORDER-RIGHT: #CCCCCC 1px solid; BORDER-BOTTOM: #CCCCCC 1px solid;BORDER-TOP: #CCCCCC 1px solid"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td background="images/host-2.gif"><img src="images/host-simple.gif" width="153" height="25"></td>
                                  </tr>
                                </table>
                                <br> 
<%
sql="select * from sample_area order by id desc" 
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
rs.pagesize=10
rs.AbsolutePage=pagecount

do while not rs.eof
memo = replace(rs("memo"),chr(13),"<br>")
%>

                                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td align="center" valign="top"><TABLE bgColor=#cccccc border=0 
                              cellPadding=5 cellSpacing=1 width="488">
                                        <TBODY>
                                          <TR> 
                                            <TD width="126" rowspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="75%" border="0" cellspacing="0" cellpadding="0">
                                                <tr>
                                                  <td height="2"><img src="1.gif" width="1" height="1"></td>
                                                </tr>
                                              </table>
                                              <a href="http://<%=rs("addr")%>" target="_blank"><img src=images/simple_host/<%=rs("img")%> border="0" width="120" height="87"></a></TD>
                                            <TD height="18" bgcolor="#f5f5f5"><b>站 名：</b> <a href="http://<%=rs("addr")%>" target="_blank"> 
                                              <%=rs("sample_area_name")%> </a></TD>
                                          </TR>
                                          <TR bgColor=#ffffff vAlign=top> 
                                            <TD height=46><b>简&nbsp;&nbsp;介：</b><br> 
                                              <br> 
                                              <% if len(rs("memo"))>99 then%>
                                              <%=left(rs("memo"),98)%>.. 
                                              <%else%>
                                              <%=rs("memo")%> 
                                              <%end if%>
                                            </TD>
                                          </TR>
                                        </TBODY>
                                      </TABLE></td>
                                  </tr>
                                </table>
                                <br> <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td height="1" bgcolor="#CCCCCC"><img src="1.gif" width="1" height="1"></td>
                                  </tr>
                                </table><br>
                                <%
rs.movenext
n=n+1                                                                     
if n>=rs.pagesize then exit do                                                           
loop
%>
                                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td>&nbsp;</td>
                                  </tr>
                                </table> 
                                <table width="488" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#f5f5f5">
                                          <tr> 
                                            <td width="50%" height="20">
											
<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr><td></td><td>            
<%
zj=1
for zi=1 to rs.pagecount
%> 
   <a href="virtual-hosting-service-Sample.asp?page=<%=zj%>">|&nbsp;<%=zj%>&nbsp;|</a>         
<%
if zj mod 12 =0 then response.Write("<br><br>")
zj=zj+1
next
%> 
</td></tr>
</table> 
									
											
											
											</td>
                                            <td width="50%" height="20" align="right">
											
											
											
											
                                        共&nbsp;<%=amount%>&nbsp;个&nbsp;&nbsp;每页&nbsp;10&nbsp;个&nbsp;&nbsp;<%=pagecount%></font>/<%=rs.pagecount%> 
                                        <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%>
                                        <a href="virtual-hosting-service-Sample.asp?page=<%=cstr(pagecount+1)%>"><font color="#ff0000">下一页</font></a><font color="#ff0000"> 
                                        <% end if %>
                                        <% if rs.pagecount>1 and rs.pagecount=pagecount then %>
                                        <a href="virtual-hosting-service-Sample.asp?page=<%=cstr(pagecount-1)%>"> 
                                        上一页</a> 
                                        <%end if%>
                                        <% if pagecount<>1 and rs.pagecount<>pagecount then%>
                                        <a href="virtual-hosting-service-Sample.asp?page=<%=cstr(pagecount-1)%>">上一页</a> 
                                        <a href="virtual-hosting-service-Sample.asp?page=<%=cstr(pagecount+1)%>"><font color="#ff0000">下一页</font></a><font color="#ff0000"> 
                                        <% end if%>											</td>
                                          </tr>
                                        </table>
                                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                  <tr>
                                    <td>&nbsp;</td>
                                  </tr>
                                </table> 
                              </TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <br> </td>
                    </tr>
                  </table>
<%
rs.close
set rs=nothing
%>
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
