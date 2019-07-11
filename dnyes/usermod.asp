<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%on error resume next
%>
<%
if session("y_it_uid")="" then
response.redirect "error.asp?err=014&zurl=usermod.asp"
response.end
end if
%>
<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from user where uid='"&session("y_it_uid")&"'"
rs.open sql,conn,1,1
%>
<script language="JavaScript">
function CheckForm()
{

if (document.modiinfo.email.value.length == 0) {
		alert("请输入您的Email.");
		document.modiinfo.email.focus();
		return false;
	}
if (document.modiinfo.email.value.length > 0 && !document.modiinfo.email.value.match( /^.+@.+$/ ) ) {
		alert("Email 错误！请重新输入");
		document.modiinfo.email.focus();
		return false;
	}
if (document.modiinfo.namez.value.length == 0) {
		alert("请输入企业/单位的中文姓名.");
		document.modiinfo.namez.focus();
		return false;
	}
if (document.modiinfo.contactz.value.length == 0) {
		alert("请输入联系人的中文姓名.");
		document.modiinfo.contactz.focus();
		return false;
	}
if (document.modiinfo.govz.value.length == 0) {
		alert("请输入中文国家名.");
		document.modiinfo.govz.focus();
		return false;
	}
if (document.modiinfo.shengz.value.length == 0) {
		alert("请输入中文省份名.");
		document.modiinfo.shengz.focus();
		return false;
	}
	if (document.modiinfo.cityz.value.length == 0) {
		alert("请输入中文城市名.");
		document.modiinfo.cityz.focus();
		return false;
	}
if (document.modiinfo.dizhiz.value.length == 0) {
		alert("请输入中文联系地址.");
		document.modiinfo.dizhiz.focus();
		return false;
	}
if (document.modiinfo.postnumber.value.length == 0) {
		alert("请输入您的邮政编码.");
		document.modiinfo.postnumber.focus();
		return false;
	}
	if (document.modiinfo.tel.value.length == 0) {
		alert("请输入您的联系电话，以便我们可以为您更好服务.");
		document.modiinfo.tel.focus();
		return false;
	}
	
	return true;
}
</script>
<html>
<head>
<title><%=Application("y_it_itname")%>-修改资料</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css.css" type="text/css">
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
                          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                            <form name="modiinfo" method="POST" action="usermodok.asp" onSubmit="return CheckForm();">
                              <tr> 
                                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
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
                                      </td>
                                    </tr>
                                    <tr> 
                                      <td width="70%"> <table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                                          <tr bgcolor="#f5f5f5"> 
                                            <td height="35" colspan="2" bgcolor="#f5f5f5"><font color="#FF0033">请您认真填写下列资料</font></td>
                                            <td height="35" align="right">带<b><font color="#FF3333">*</font></b>的必须填写</td>
                                          </tr>
                                          <tr> 
                                            <td width="145" height="30" bgcolor="#f5f5f5">用户帐号：</td>
                                            <td width="174" bgcolor="#FFFFFF"><%=rs("uid")%></td>
                                            <td width="145" bgcolor="#FFFFFF"></td>
                                          </tr>
                                          <tr> 
                                            <td height="25" bgcolor="#f5f5f5">电子邮件：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="email" size="20" maxlength="40" class="form" value=<%=rs("email")%>> 
                                            </td>
                                            <td bgcolor="#FFFFFF"><b><font color="#FF3333">*</font></b></td>
                                          </tr>
                                          <tr> 
                                            <td height="25" bgcolor="#f5f5f5">性&nbsp;&nbsp;&nbsp;&nbsp;别：</td>
                                            <td bgcolor="#FFFFFF"> <select class=form name=sex >
                                                <option selected  value="<%=rs("sex")%>"><%=rs("sex")%></option>
                                                <option  value="男">男</option>
                                                <option value="女">女</option>
                                              </select> </td>
                                            <td bgcolor="#FFFFFF"><b><font color="#FF3333">*</font></b></td>
                                          </tr>
                                          <tr> 
                                            <td height="25" bgcolor="#f5f5f5">年&nbsp;&nbsp;&nbsp;&nbsp;龄：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="age" size="3" maxlength="3" class="form" value="<%=rs("age")%>"> 
                                            </td>
                                            <td bgcolor="#FFFFFF"></td>
                                          </tr>
                                          <tr> 
                                            <td height="25" bgcolor="#f5f5f5">企业/单位（中文）：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="namez" size="20" maxlength="80" class="form" value="<%=rs("namez")%>"> 
                                            </td>
                                            <td bgcolor="#FFFFFF"><b><font color="#FF3333">*</font></b></td>
                                          </tr>
                                          <tr> 
                                            <td height="25" bgcolor="#f5f5f5">企业/单位（英文）：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="namee"  maxlength="120" size="20" class="form" value="<%=rs("namee")%>"> 
                                            </td>
                                            <td bgcolor="#FFFFFF"></td>
                                          </tr>
                                          <tr> 
                                            <td height="25" bgcolor="#f5f5f5">联系人（中文）：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="contactz" size="20" maxlength="20" class="form" value="<%=rs("contactz")%>"> 
                                            </td>
                                            <td bgcolor="#FFFFFF"><b><font color="#FF3333">*</font></b></td>
                                          </tr>
                                          <tr> 
                                            <td height="25" bgcolor="#f5f5f5">联系人（英文）：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="contacte"  maxlength="30" size="20" class="form" value="<%=rs("contacte")%>"> 
                                            </td>
                                            <td bgcolor="#FFFFFF"></td>
                                          </tr>
                                          <tr> 
                                            <td height="25" bgcolor="#f5f5f5">国家（中文）：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="govz" size="20" maxlength="30" class="form" value="<%=rs("govz")%>"> 
                                            </td>
                                            <td bgcolor="#FFFFFF"><b><font color="#FF3333">*</font></b></td>
                                          </tr>
                                          <tr> 
                                            <td height="25" bgcolor="#f5f5f5"> 国家（英文）：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="gove"  maxlength="50" size="20" class="form" value=<%=rs("gove")%>> 
                                            </td>
                                            <td bgcolor="#FFFFFF"></td>
                                          </tr>
                                          <tr> 
                                            <td height="25" bgcolor="#f5f5f5"> 省份（中文）：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="shengz" size="20" maxlength="30" class="form" value=<%=rs("shengz")%>> 
                                            </td>
                                            <td bgcolor="#FFFFFF"><b><font color="#FF3333">*</font></b></td>
                                          </tr>
                                          <tr> 
                                            <td height="25" bgcolor="#f5f5f5">省份（英文）：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="shenge"  maxlength="50" size="20" class="form" value=<%=rs("shenge")%>> 
                                            </td>
                                            <td bgcolor="#FFFFFF"></td>
                                          </tr>
                                          <tr> 
                                            <td bgcolor="#f5f5f5">城市（中文）：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="cityz" size="20" maxlength="30" class="form" value=<%=rs("cityz")%>> 
                                            </td>
                                            <td bgcolor="#FFFFFF"><b><font color="#FF3333">*</font></b></td>
                                          </tr>
                                          <tr> 
                                            <td bgcolor="#f5f5f5">城市（英文）：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="citye"  maxlength="50" size="20" class="form" value=<%=rs("citye")%>> 
                                            </td>
                                            <td bgcolor="#FFFFFF"></td>
                                          </tr>
                                          <tr> 
                                            <td bgcolor="#f5f5f5">&nbsp地址（中文）：;</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="dizhiz" size="20" maxlength="80" class="form" value=<%=rs("dizhiz")%>> 
                                            </td>
                                            <td bgcolor="#FFFFFF"><b><font color="#FF3333">*</font></b></td>
                                          </tr>
                                          <tr> 
                                            <td bgcolor="#f5f5f5">地址（英文）：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="dizhie"  maxlength="120" size="20" class="form" value=<%=rs("dizhie")%>> 
                                            </td>
                                            <td bgcolor="#FFFFFF"></td>
                                          </tr>
                                          <tr> 
                                            <td bgcolor="#f5f5f5">邮政编码：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="postnumber" size="20" maxlength="8" class="form" value=<%=rs("postnumber")%>> 
                                            </td>
                                            <td bgcolor="#FFFFFF"><b><font color="#FF3333">*</font></b></td>
                                          </tr>
                                          <tr> 
                                            <td bgcolor="#f5f5f5">联系电话-1：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="tel"  maxlength="30" size="20" class="form" value=<%=rs("tel")%>> 
                                            </td>
                                            <td bgcolor="#FFFFFF"><b><font color="#FF3333">*</font></b></td>
                                          </tr>
                                          <tr> 
                                            <td bgcolor="#f5f5f5">联系电话-2：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="tel2"  maxlength="30" size="20" class="form" value=<%=rs("tel2")%>> 
                                            </td>
                                            <td bgcolor="#FFFFFF"></td>
                                          </tr>
                                          <tr> 
                                            <td bgcolor="#f5f5f5">OICQ/ICQ/MSN：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="oicq" size="20" maxlength="40" class="form" value=<%=rs("oicq")%>> 
                                            </td>
                                            <td bgcolor="#FFFFFF"></td>
                                          </tr>
                                          <tr> 
                                            <td bgcolor="#f5f5f5">传真：</td>
                                            <td bgcolor="#FFFFFF"> <input type="text" name="fax" size="20" maxlength="30" class="form" value=<%=rs("fax")%>> 
                                            </td>
                                            <td bgcolor="#FFFFFF"></td>
                                          </tr>
                                          <tr> 
                                            <td bgcolor="#f5f5f5">&nbsp;</td>
                                            <td bgcolor="#FFFFFF"></td>
                                            <td bgcolor="#FFFFFF"></td>
                                          </tr>
                                          <tr bgcolor="#FFFFFF"> 
                                            <td colspan="3">&nbsp;</td>
                                          </tr>
                                          <tr bgcolor="#f5f5f5"> 
                                            <td height="35" colspan="3"><div align="center"> 
                                                <input class=botton type="submit" name="ok" value="修 改">
                                                &nbsp;&nbsp;&nbsp; 
                                                <input class=botton type="reset" name="Reset" value="清 除">
                                                &nbsp;&nbsp;&nbsp; 
                                                <input type="hidden" name="modiinfo" value="true">
                                              </div></td>
                                          </tr>
                                        </table></td>
                                    </tr>
                                    <tr> 
                                      <td>&nbsp;</td>
                                    </tr>
                                  </table></td>
                              </tr>
                              <tr> 
                                <td colspan="2"><div align="center"></div></td>
                              </tr>
                            </form>
                            <%
rs.close
set rs=nothing
%>
                          </table>
                          <br>
                          <br>
                        </TD>
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