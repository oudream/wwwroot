<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
<%on error resume next%>
<%
id=request("id")
sql="select * from subs where id="&id 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
id=rs("id")
photo=rs("photo")
area=rs("area")
price=rs("price")
subs=rs("subs")
subsname=rs("subsname")
bookbm=rs("bookbm")
subsunit=rs("subsunit")
gys=rs("gys")

if rs("ifdisc")=true then ifdisc="true"
other = replace(rs("other"),chr(13),"<br>")
brief = replace(rs("brief"),chr(13),"<br>")
rs.close
set rs=nothing
sql="select * from area where area='"&area&"'"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
introf = replace(rs("introf"),chr(13),"<br>")
rs.close
set rs=nothing
%>
<script language="JavaScript">
function MM_jumpMenu(selObj,restore){ 
  eval("window.open('"+selObj.options[selObj.selectedIndex].value+"')");
  if (restore) selObj.selectedIndex=0;
}
</script>
<html>
<head>
<title><%=Application("y_it_itname")%>-<%=area%>-<%=subsname%></title>
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
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>客 户 帮 助</b></font></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td> <TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
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
                      <td valign="top">
                        <!--#include file="faqhost.asp" -->
                      </td>
                      <td width="2" valign="top" background="images/left-2-right.gif"><img src="images/left-2-right.gif" width="2" height="5"></td>
                    </tr>
                  </table>
                  <TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
                    <TBODY>
                      <TR> 
                        <TD height="2" 
                vAlign=top 
                style="BORDER-left: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                      </TR>
                    </TBODY>
                  </TABLE></td>
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
                <td valign="top"> <!--#include file="support.asp" --> </td>
              </tr>
            </table></td>
          <td width="7">&nbsp;</td>
          <td width="513" valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="8"></td>
              </tr>
              <tr> 
                <td height="25"><p
                         align="center" style="margin-top:3px;"> 
                  <div align="center"><b>=== <%=subsname%> ===</b> <br>
                    <br>
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td bgcolor="#f0f0f0"><img src="1.gif" width="1" height="1"></td>
                      </tr>
                    </table>
                    <br>
                  </div></td>
              </tr>
            </table>
            <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td>
                  <table width="500" border="0" cellspacing="0" cellpadding="0" align="center">
                    <tr> 
                      <td width="100" height="120" rowspan="4" valign="top"><a href="add.asp?add=<%=bookbm%>" target="_blank"><img src=<%=photo%> width="100" height="120" border="0"></a></td>
                      <td width="6" rowspan="4" align="center" valign="top">&nbsp;</td>
                      <td height="20" width="200"><b>编号：</b> <%=id%></td>
                      <td width="194"><font color=#666666>[服务商： <%=gys%>]</font></td>
                    </tr>
                    <tr> 
                      <td height="20"><b>名称：</b> <%=subsname%></td>
                      <td height="20"><font color=#666666>[类&nbsp;&nbsp;别： <%=area%>]</font></td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2"><b>价格：</b> 
                        <%if ifdisc="true" then%>
                        <%response.write FormatNumber(csng(price)*session("y_it_userdiscount"),2)%>
                        <%else%>
                        <%response.write FormatNumber(csng(price),2)%>
                        <%end if%>
                        RMB /<%=subsunit%></td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td colspan="6" valign="top"><br><%=brief%><br><br><%=other%><br></td>
                    </tr>
                    <tr> 
                      <td colspan="6" height="35"> <div align="center"> 
                          <table width="500" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td width="250" height="25"> <div align="right"><a href="add.asp?add=<%=bookbm%>" target="_blank"><img src="images/in3.gif" border="0"></a></div></td>
                              <form method="post">
                                <td> <div align="right"> 
                                    <select name="id" onchange="MM_jumpMenu(this,0)">
                                      <option value="" selected><%=area%>其他商品</option>
<%                           
sql="select * from subs where area='"&area&"' and subs<>'"&subs&"'"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
while not rs.eof
%>
                                      <option value="showsubs.asp?id=<%=rs("id")%>"><%=rs("subsname")%></option>
<%
rs.movenext
wend
rs.Close()
%>
                                    </select>
                                  </div></td>
                              </form>
                            </tr>
                          </table>
                        </div></td>
                    </tr>
                  </table>
                  <br> </td>
              </tr>
              <tr> 
                <td height="20">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="2">&nbsp;</td>
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
