<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
<%on error resume next%>
<%
if session("y_it_uid")="" then
response.redirect "error.asp?err=014&zurl=viewmyordersb.asp"
response.end
end if
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-查看我的定单</title>
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
            <TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
              <TBODY>
                <TR> 
                  <TD height="3" 
                vAlign=top 
                style="BORDER-left: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                </TR>
              </TBODY>
            </TABLE>
            <table width="228" height="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="1" background="images/1x2-black.gif"><img src="images/1x2-black.gif" width="1" height="2"></td>
                <td height="180" valign="top"> 
                  <!--#include file="support.asp" -->
                </td>
              </tr>
            </table></td>
          <td width="7">&nbsp;</td>
          <td width="513" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr> 
                <td colspan="5" height="88">
<TABLE cellSpacing=0 cellPadding=0 width=100% border=0>
                    <TBODY>
                      <TR> 
                        <TD width="34%" vAlign=center style="BORDER-left: #CCCCCC 1px solid; BORDER-RIGHT: #CCCCCC 1px solid; BORDER-BOTTOM: #CCCCCC 1px solid;BORDER-TOP: #CCCCCC 1px solid"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td width="12" height="25" background="images/host-2.gif">&nbsp;</td>
                              <td valign="bottom" background="images/host-2.gif">查 
                                看 定 单</td>
                            </tr>
                          </table>
                          <table width="75%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td height="6"><img src="1.gif" width="1" height="1"></td>
                            </tr>
                          </table>
                          <table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#cccccc">
                            <tr bgcolor="#f5f5f5"> 
                              <td width="28%" height="35">以下是您和我们交易的所有定单，您可以方便查询其处理情况和定单所有详细信息</td>
                            </tr>
                          </table>
                          <br>
                          <table width="498" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
          <tr bgcolor="#9CCCFF"> 
            <td width="55" height="35" align="center" bgcolor="#f5f5f5">ID</td>
            <td width="65" align="center" bgcolor="#f5f5f5">订单号</td>
            <td width="58" align="center" bgcolor="#f5f5f5">支付金额</td>
            <td width="58" align="center" bgcolor="#f5f5f5">支付方式</td>
            <td width="108" align="center" bgcolor="#f5f5f5">订购日期</td>
            <td width="45" align="center" bgcolor="#f5f5f5">详情</td>
			<td width="45" align="center" bgcolor="#f5f5f5">删除</td>
          </tr>
<%
sql="select * from orders where username='"&session("y_it_uid")&"' order by id desc"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1 

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1   
end if

rs.pagesize=20
rs.AbsolutePage=pagecount
%>
<% do while not rs.eof %> 
          <tr bgcolor="#FFFFFF"> 
            <td height="25" align="center" bgcolor="#f5f5f5"><%=rs("id")%></td>
            <td align="center"><%=rs("inBillNo")%></td>
            <td align="center"><%=rs("summoney")%></td>
            <td align="center"><%=rs("paymenttype")%></td>
            <td align="center"><%=rs("ordertime")%></td>
            <td align="center">[<a href="viewmyorders-1.asp?id=<%=rs("id")%>">详情</a>]</td>
            <td align="center">[<a href="ifdelorder.asp?id=<%=rs("id")%>" >删除</a>]</td>
          </tr>
          <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
%> 
          <tr bgcolor="#FFFFFF"> 
            <form action="viewmyorders.asp" method="post">
              <td height="35" colspan="9"> 
                <div align="center"> 共 <b><%=rs.recordcount%></b> 张订单, 页次: <b><font color=red><%=pagecount%></font>/<%=rs.pagecount%></b>, 
                  当前从第 <%
if pagecount<=1 then
response.write "<font color=red>1</font>"
else
response.write "<font color=red>" & pagecount*rs.pagesize-rs.pagesize+1 & "</font>"
end if
%> 张开始。 <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%> 
                  <a href="viewmyorders.asp?page=<%=cstr(pagecount+1)%>"> 下一页</a> 
                  <% end if %> <% if rs.pagecount>1 and rs.pagecount=pagecount then %> 
                  <a href="viewmyorders.asp?page=<%=cstr(pagecount-1)%>"> 上一页</a> 
                  <%end if%> <% if pagecount<>1 and rs.pagecount<>pagecount then%> 
                  <a href="viewmyorders.asp?page=<%=cstr(pagecount-1)%>"> 上一页</a> 
                  <a href="viewmyorders.asp?page=<%=cstr(pagecount+1)%>"> 下一页</a> 
                  <% end if%>&nbsp; 直接到第 
                  <select name="page">
                    <%for i=1 to rs.pagecount%> 
                    <option value="<%=i%>"><%=i%></option>
                    <%next%> 
                  </select>
                  页 
                  <input class=botton type="submit" name="go" value="&lt; GO &gt;">
                  <input type="hidden" name="uid" value=<%=uid%>>
                  <input type="hidden" name="viewcomp" value=<%=viewcomp%>>
                </div>
              </td>
            </form>
          </tr>
        </table>
        <%
rs.close
set rs=nothing
conn.close
%>
                          <table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td>&nbsp;</td>
                            </tr>
                          </table>
                          <table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td>&nbsp;</td>
                            </tr>
                          </table></TD>
                      </TR>
                    </TBODY>
                  </TABLE></td>
              </tr>
            </table></td>
          <td width="7" background="images/right-7x5.gif">&nbsp;</td>
        </tr>
      </table></td>
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
  </tr>
</table>
<!--#include file="copyright.asp"-->
</body>
</html>
