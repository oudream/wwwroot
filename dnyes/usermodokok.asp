<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
<%on error resume next%>
<%
if session("y_it_uid")="" then
response.redirect "index.asp"
response.end
end if
%>
<%
'从提交表单返回值，并查看是否有非法自符

sql = "SELECT * FROM user where uid= '" &session("y_it_uid")& "'"
set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
email=rs("email")
sex=rs("sex")
age=rs("age")
namez=rs("namez")
namee=rs("namee")
contactz=rs("contactz")
contacte=rs("contacte")
govz=rs("govz")
gove=rs("gove")
shengz=rs("shengz")
shenge=rs("shenge")
cityz=rs("cityz")
citye=rs("citye")
dizhiz=rs("dizhiz")
dizhie=rs("dizhie")
tel=rs("tel")
tel2=rs("tel2")
oicq=rs("oicq")
postnumber=rs("postnumber")
fax=rs("fax")

if age="" then age=" "
if namee="" then namee=" "
if contacte="" then contacte=" "
if gove="" then gove=" "
if shenge="" then shenge=" "
if citye="" then citye=" "
if dizhie="" then dizhie=" "
if tel2="" then tel2=" "
if fax="" then fax=" "
if oicq="" then oicq=" "


conn.execute "update user set email='"&email&"',sex='"&sex&"',age='"&age&"',namez='"&namez&"',namee='"&namee&"',contactz='"&contactz&"',contacte='"&contacte&"',govz='"&govz&"',gove='"&gove&"',shengz='"&shengz&"',shenge='"&shenge&"',cityz='"&cityz&"',citye='"&citye&"',dizhiz='"&dizhiz&"',dizhie='"&dizhie&"',tel='"&tel&"',tel2='"&tel2&"',postnumber='"&postnumber&"',oicq='"&oicq&"',fax='"&fax&"' where uid='"&session("y_it_uid")&"'"
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-修改资料</title>
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
                <td colspan="5" height="88"> <TABLE cellSpacing=0 cellPadding=0 width=100% border=0>
                    <TBODY>
                      <TR> 
                        <TD width="34%" vAlign=center style="BORDER-left: #CCCCCC 1px solid; BORDER-RIGHT: #CCCCCC 1px solid; BORDER-BOTTOM: #CCCCCC 1px solid;BORDER-TOP: #CCCCCC 1px solid"> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td height="8" colspan="2"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td width="12" height="25" background="images/host-2.gif">&nbsp;</td>
                                    <td valign="bottom" background="images/host-2.gif">修 
                                      改 资 料</td>
                                  </tr>
                                </table>
                                <table width="75%" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td height="6"><img src="1.gif" width="1" height="1"></td>
                                  </tr>
                                </table>
                                <table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                                  <tr bgcolor="#f5f5f5"> 
                                    <td height="35" colspan="2">您已经成功的修改了您的个人资料,<font color="#FF0033">以下是您的新资料</font>。</td>
                                  </tr>
                                  <tr> 
                                    <td width="124" height="25" bgcolor="#f5f5f5">用户帐号：</td>
                                    <td width="351" bgcolor="#FFFFFF"><%=session("y_it_uid")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">电子邮件：</td>
                                    <td bgcolor="#FFFFFF"><%=email%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">企业/单位（中文）：</td>
                                    <td bgcolor="#FFFFFF"><%=namez%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">企业/单位（英文）： 
                                    </td>
                                    <td bgcolor="#FFFFFF"><%=namee%> </td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">联系人（中文）：</td>
                                    <td bgcolor="#FFFFFF"><%=contactz%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">联系人（英文）： 
                                    </td>
                                    <td bgcolor="#FFFFFF"><%=contacte%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">性&nbsp;&nbsp;&nbsp;&nbsp;别：</td>
                                    <td bgcolor="#FFFFFF"><%=sex%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">年&nbsp;&nbsp;&nbsp;&nbsp;龄：</td>
                                    <td bgcolor="#FFFFFF"><%=age%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">国&nbsp;&nbsp;&nbsp;&nbsp;家：</td>
                                    <td bgcolor="#FFFFFF"><%=govz%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">省份（中文）： 
                                    </td>
                                    <td bgcolor="#FFFFFF"><%=shengz%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">省份（英文）： 
                                    </td>
                                    <td bgcolor="#FFFFFF"><%=shenge%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">城市（中文）：</td>
                                    <td bgcolor="#FFFFFF"><%=cityz%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">城市（英文）： 
                                    </td>
                                    <td bgcolor="#FFFFFF"><%=citye%></td>
                                  </tr>
                                  <tr> 
                                    <td height="5" bgcolor="#f5f5f5">地址（中文）：</td>
                                    <td bgcolor="#FFFFFF"><%=dizhiz%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">地址（英文）： 
                                    </td>
                                    <td bgcolor="#FFFFFF"><%=dizhie%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">邮政编码：</td>
                                    <td bgcolor="#FFFFFF"><%=postnumber%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">联系电话-1：</td>
                                    <td bgcolor="#FFFFFF"><%=tel%></td>
                                  </tr>
                                  <tr> 
                                    <td height="5" bgcolor="#f5f5f5">联系电话-2：</td>
                                    <td bgcolor="#FFFFFF"><%=tel2%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">传&nbsp;&nbsp;&nbsp;&nbsp;真：</td>
                                    <td bgcolor="#FFFFFF"><%=fax%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">OICQ/ICQ/MSN：</td>
                                    <td bgcolor="#FFFFFF"><%=oicq%></td>
                                  </tr>
                                  <tr bgcolor="#FFFFFF"> 
                                    <td height="50" colspan="2">&nbsp;</td>
                                  </tr>
                                  <tr bgcolor="#f5f5f5"> 
                                    <td height="40" colspan="2">如果您要更改了您的密码，请<a href="usermodpass.asp"><font color="#FF0000">点此</font></a>进行操作</td>
                                  </tr>
                                </table></td>
                            </tr>
                            <tr> 
                              <td width="70%">&nbsp; </td>
                            </tr>
                            <tr> 
                              <td>&nbsp;</td>
                            </tr>
                          </table>
                          <br> <br> </TD>
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
<%
rs.close
set rs=nothing
%>  
<!--#include file="copyright.asp"-->
</body>
</html>
