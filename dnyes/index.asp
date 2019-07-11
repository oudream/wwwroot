<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%on error resume next%>
<%
sqlsys="select * from sysdefault" 
Set rssys=Server.CreateObject("ADODB.RecordSet") 
rssys.Open sqlsys,conn,1,1 
Application("y_it_fromemail")=rssys("fromemail")
Application("y_it_url")=rssys("url")
Application("y_it_itname")=rssys("esname")
rssys.close
set rssys=nothing
%>
<html>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".open('"+selObj.options[selObj.selectedIndex].value+"')");
  if (restore) selObj.selectedIndex=0;
}

function doValidate()
{
	if(form1.domain.value.length == 0) {
		alert("请填写一个域名！");
		form1.domain.focus();
		return false;
	}
	if(form1.domain.value.length > 26) {
		alert("域名超长！请不要超过26个，否则不能注册！");
		form1.domain.focus();
		return false;
	}
	if(! isDomain(form1.domain.value)) {
		alert("域名不合法！\n域名不能以中横杠'－'开头或结尾！\n除英文26个字母和10个阿拉伯数字以及中横杠'－'可以用作域名，其它的，请不要作域名！");
		form1.domain.focus();
		return false;
	}
	return true;
}

function isDomain(name)
{
	if(name.length > 0) {
		if(name.charAt(0) == "-" || name.charAt(name.length - 1) == "-")
			return false;
	}
	for(i = 0; i < name.length; i++) {
		if((name.charAt(i) < "a" || name.charAt(i) > "z") && (name.charAt(i) < "A" || name.charAt(i) > "Z") && (name.charAt(i) < "0" || name.charAt(i) > "9") && name.charAt(i) != "-")
			return false;
	}
	return true;
}

//-->
</script>
<head>
<title>信网 数信科技-佛山 顺德 网站建设 网页制作 网站推广 域名注册 虚拟主机 企业邮箱 租用托管</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css.css" type="text/css">
</head>
<body leftmargin="0" topmargin="0">
<!--#include file="top.asp"-->
<table width="776" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
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
          <td width="229" valign="top" background="images/left-228x5.gif"><table width="229" border="0" cellpadding="0" cellspacing="0">
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
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="228" height="28" background="images/left-2.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="10%" height="18" valign="bottom">&nbsp; </td>
                      <td width="8%" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>相 
                        关 知 识</b></font></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td>
                  <TABLE width=100% border=0 cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
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
                      <td height="100" valign="top">
					  <!--#include file="base-index.asp" -->
					  </td>
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
                  </TABLE>
				  </td>
              </tr>
            </table>
            <table width="229" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="229" height="28" background="images/left-2.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="10%" height="18" valign="bottom">&nbsp; </td>
                      <td width="8%" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>常 
                        用 文 档</b></font></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="1" background="images/1x2-black.gif"><img src="images/1x2-black.gif" width="1" height="2"></td>
                      <td>
                  <TABLE width=100% border=0 cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
                    <TBODY>
                      <TR> 
                        <TD height="2"><img src="1.gif" width="1" height="1"></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td><!--#include file="document-index.asp" --></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="7">&nbsp;</td>
          <td width="305" valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td colspan="2" align="center"><table width="75%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td height="1"><img src="1.gif" width="1" height="1"></td>
                    </tr>
                  </table>
                  <table width="295" height="135" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td valign="top" background="images/middle-1.gif"><form name="form1" id="form1" action="domain.asp" method="post" onSubmit="return doValidate();">
                          <table width="100%" border="0" cellpadding="0" cellspacing="3">
                            <tr> 
                              <td height="30" colspan="4">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td width="70">&nbsp;</td>
                              <td width="32" class=eng valign="bottom">www.</td>
                              <td width="103"><font color="#000000"> 
                                <input class="input" type="text" name="domain" id=domain maxlength="20" size="16">
                                <input type="hidden" name="domaintypestr" value="国际域名查询">
                                </font></td>
                              <td width="75"><font color="#000000"> 
                                <input type="submit" class=botton name=Submit  id=Submit  value="查 询" align="middle" border="0" width="80" height="25">
                                </font></td>
                            </tr>
                            <tr> 
                              <td height="20" class=eng colspan="4"> <input type="radio" name="after" value=".com" checked>
                                .com 
                                <input type="radio" name="after" value=".net">
                                .net 
                                <input type="radio" name="after" value=".org">
                                .org 
                                <input type="radio" name="after" value=".biz">
                                .biz 
                                <input type="radio" name="after" value=".info">
                                .info</td>
                            </tr>
                            <tr> 
                              <td height="25" class=eng colspan="4"> <input type="radio" name="after" value=".com.cn">
                                .com.cn 
                                <input type="radio" name="after" value=".net.cn">
                                .net.cn 
                                <input type="radio" name="after" value=".org.cn">
                                .org.cn 
                                <input type="radio" name="after" value=".cn"> 
                                <font color="#FF0000"> .cn </font></td>
                            </tr>
                          </table>
                        </form></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td height="10" colspan="2">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="2" background="images/middle-2.gif"><img src="images/middle-2.gif" width="5" height="2"></td>
              </tr>
              <tr> 
                <td height="186" colspan="2"><TABLE cellSpacing=1 cellPadding=1 width="100%" border=0>
                    <TBODY>
                      <TR> 
                        <TD vAlign=top bgColor=#ffffff rowSpan=3><a 
                  href="showsubs.asp?id=204" target="_blank"><img src="products/index/index_1.gif" width="80" height="100" 
                  hspace=0 border=0></a></TD>
                        <TD align=right height=22> <TABLE cellSpacing=1 cellPadding=0 width="98%" bgColor=#f2f2f2 
                  border=0>
                            <TBODY>
                              <TR> 
                                <TD><FONT color=#000000>&nbsp;</FONT><FONT color=#ffffff><FONT 
                        color=#ffffff><IMG height=9 src="images/dot.gif" 
                        width=4></FONT></FONT><FONT color=#000000><IMG height=12 
                        src="images/pic111.gif" width=12> <B>商业标准型I-ASP </B></FONT></TD>
                                <TD width=34 bgColor=#f3f3f3><FONT color=#000000><A 
                        href="showsubs.asp?id=204" target="_blank"><FONT 
                        color=#666666><IMG height=20 src="images/B.GIF" 
                        width=32 border=0></FONT></A></FONT></TD>
                                <TD width=34 bgColor=#f3f3f3> <DIV align=center><FONT color=#000000><A 
                        href="http://dns99.cn/detail.php?gid=116&amp;nowcatid=6"></A><A 
                        href="add.asp?add=204"><FONT 
                        color=#666666><IMG height=20 src="images/M.GIF" 
                        width=32 
                  border=0></FONT></A></FONT></DIV></TD>
                              </TR>
                            </TBODY>
                          </TABLE></TD>
                      </TR>
                      <TR> 
                        <TD align=right height=22> <TABLE cellSpacing=1 cellPadding=0 width="98%" border=0>
                            <TBODY>
                              <TR> 
                                <TD width="55%" height=21><FONT color=#000000>价格：<FONT 
                        color=#999999><FONT color=#000000>195.00元/<FONT 
                        color=#000000><FONT color=#999999><FONT 
                        color=#000000><FONT color=#999999><FONT 
                        color=#000000>年</FONT></FONT></FONT></FONT></FONT></FONT></FONT></FONT></TD>
                                <TD width="45%" height=21><FONT color=#000000><FONT 
                        color=#999999><FONT color=#000000>详情：</FONT> <A 
                        href="showsubs.asp?id=204" target="_blank"><FONT 
                        color=#ff0000>点此查看</FONT></A> </FONT></FONT></TD>
                              </TR>
                            </TBODY>
                          </TABLE></TD>
                      </TR>
                      <TR> 
                        <TD align=right bgColor=#ffffff> <TABLE cellSpacing=2 cellPadding=0 width="98%" border=0>
                            <TBODY>
                              <TR> 
                                <TD><FONT color=#000000>经济之选，送1个英文国际域名，60M空间，可同时处理 50 个访问请求。支持ASP+ACCESS+FLASH...</FONT></TD>
                              </TR>
                            </TBODY>
                          </TABLE>
                          <FONT 
                  color=#000000><FONT color=#999999><FONT 
                  color=#000000></FONT></FONT></FONT></TD>
                      </TR>
                      <TR> 
                        <TD vAlign=top bgColor=#ffffff colSpan=2 height=4><FONT 
                  color=#ffffff><FONT color=#ffffff><IMG height=4 
                  src="images/dot.gif" width=8></FONT></FONT></TD>
                      </TR>
                      <TR> 
                        <TD vAlign=top bgColor=#ffffff rowSpan=3><A 
                  href="showsubs.asp?id=209" target="_blank"><IMG src="products/index/index_2n3.gif" width="80" height="100" 
                  hspace=0 border=0></A></TD>
                        <TD align=right height=22> <TABLE cellSpacing=1 cellPadding=0 width="98%" bgColor=#f2f2f2 
                  border=0>
                            <TBODY>
                              <TR> 
                                <TD><FONT color=#000000>&nbsp;</FONT><FONT color=#ffffff><FONT 
                        color=#ffffff><IMG height=9 src="images/dot.gif" 
                        width=4></FONT></FONT><FONT color=#000000><IMG height=12 
                        src="images/pic111.gif" width=12> <B>商业安全型II-ASP </B></FONT></TD>
                                <TD width=34 bgColor=#f3f3f3><FONT color=#000000><A 
                        href="showsubs.asp?id=209" target="_blank"><FONT 
                        color=#666666><IMG height=20 src="images/B.GIF" 
                        width=32 border=0></FONT></A></FONT></TD>
                                <TD width=34 bgColor=#f3f3f3> <DIV align=center><FONT color=#000000><A 
                        href="http://dns99.cn/detail.php?gid=115&amp;nowcatid=6"></A><A 
                        href="add.asp?add=209"><FONT 
                        color=#666666><IMG height=20 src="images/M.GIF" 
                        width=32 
                  border=0></FONT></A></FONT></DIV></TD>
                              </TR>
                            </TBODY>
                          </TABLE></TD>
                      </TR>
                      <TR> 
                        <TD align=right height=22> <TABLE cellSpacing=1 cellPadding=0 width="98%" border=0>
                            <TBODY>
                              <TR> 
                                <TD width="55%" height=21><FONT color=#000000>价格：<FONT 
                        color=#999999><FONT color=#000000>295.00元/<FONT 
                        color=#000000><FONT color=#999999><FONT 
                        color=#000000><FONT color=#999999><FONT 
                        color=#000000>年</FONT></FONT></FONT></FONT></FONT></FONT></FONT></FONT></TD>
                                <TD width="45%" height=21><FONT color=#000000><FONT 
                        color=#999999><FONT color=#000000><FONT 
                        color=#999999><FONT color=#000000>详情</FONT></FONT>：</FONT><FONT 
                        color=#ff0000></FONT> <A 
                        href="showsubs.asp?id=209" target="_blank"><FONT 
                        color=#ff0000>点此查看</FONT></A> </FONT></FONT></TD>
                              </TR>
                            </TBODY>
                          </TABLE></TD>
                      </TR>
                      <TR> 
                        <TD align=right bgColor=#ffffff> <TABLE cellSpacing=2 cellPadding=0 width="98%" border=0>
                            <TBODY>
                              <TR> 
                                <TD><FONT color=#000000>超值之选，送1个英文国际域名，150M空间任意分配网站空间和企业邮箱空间。支持ASP+ACCESS+FLASH...</FONT></TD>
                              </TR>
                            </TBODY>
                          </TABLE>
                          <FONT 
                  color=#000000><FONT color=#999999><FONT 
                  color=#000000></FONT></FONT></FONT></TD>
                      </TR>
                      <TR> 
                        <TD vAlign=top bgColor=#ffffff colSpan=2 height=4><FONT 
                  color=#ffffff><FONT color=#ffffff><IMG height=4 
                  src="images/dot.gif" width=8></FONT></FONT></TD>
                      </TR>
                      <TR> 
                        <TD vAlign=top bgColor=#ffffff rowSpan=3><A 
                  href="showsubs.asp?id=214"><IMG src="products/index/index_2n3.gif" width="80" height="100" 
                  hspace=0 border=0></A></TD>
                        <TD align=right height=22> <TABLE cellSpacing=1 cellPadding=0 width="98%" bgColor=#f2f2f2 
                  border=0>
                            <TBODY>
                              <TR> 
                                <TD><FONT color=#000000>&nbsp;</FONT><FONT color=#ffffff><FONT 
                        color=#ffffff><IMG height=9 src="images/dot.gif" 
                        width=4></FONT></FONT><FONT color=#000000><IMG height=12 
                        src="images/pic111.gif" width=12> <B>商业安全型III-ASP </B></FONT></TD>
                                <TD width=34 bgColor=#f3f3f3><FONT color=#000000><A 
                        href="showsubs.asp?id=214" target="_blank"><FONT 
                        color=#666666><IMG height=20 src="images/B.GIF" 
                        width=32 border=0></FONT></A></FONT></TD>
                                <TD width=34 bgColor=#f3f3f3> <DIV align=center><FONT color=#000000><A 
                        href="http://dns99.cn/detail.php?gid=130&amp;nowcatid=6"></A><A 
                        href="add.asp?add=214"><FONT 
                        color=#666666><IMG height=20 src="images/M.GIF" 
                        width=32 
                  border=0></FONT></A></FONT></DIV></TD>
                              </TR>
                            </TBODY>
                          </TABLE></TD>
                      </TR>
                      <TR> 
                        <TD align=right height=22> <TABLE cellSpacing=1 cellPadding=0 width="98%" border=0>
                            <TBODY>
                              <TR> 
                                <TD width="55%" height=21><FONT color=#000000>价格：<FONT 
                        color=#999999><FONT color=#000000>395.00元/<FONT 
                        color=#000000><FONT color=#999999><FONT 
                        color=#000000><FONT color=#999999><FONT 
                        color=#000000>年</FONT></FONT></FONT></FONT></FONT></FONT></FONT></FONT></TD>
                                <TD width="45%" height=21><FONT color=#000000><FONT 
                        color=#999999><FONT color=#000000><FONT 
                        color=#999999><FONT color=#000000>详情</FONT></FONT>：</FONT><FONT 
                        color=#ff0000></FONT> <A 
                        href="showsubs.asp?id=214" target="_blank"><FONT 
                        color=#ff0000>点此查看</FONT></A> </FONT></FONT></TD>
                              </TR>
                            </TBODY>
                          </TABLE></TD>
                      </TR>
                      <TR> 
                        <TD align=right bgColor=#ffffff> <TABLE cellSpacing=2 cellPadding=0 width="98%" border=0>
                            <TBODY>
                              <TR> 
                                <TD><FONT 
                        color=#000000>实惠之选，送1个英文国际域名，220M空间任意分配网站空间和企业邮箱空间。支持ASP+ACCESS+FLASH...</FONT></TD>
                              </TR>
                            </TBODY>
                          </TABLE>
                          <FONT 
                  color=#000000><FONT color=#999999><FONT 
                  color=#000000></FONT></FONT></FONT></TD>
                      </TR>
                      <TR> 
                        <TD vAlign=top bgColor=#ffffff colSpan=2 height=4><FONT 
                  color=#ffffff><FONT color=#ffffff><IMG height=4 
                  src="images/dot.gif" width=8></FONT></FONT></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  <TABLE cellSpacing=1 cellPadding=1 width="100%" border=0>
                    <TBODY>
                      <TR> 
                        <TD vAlign=top bgColor=#ffffff rowSpan=3><a 
                  href="showsubs.asp?id=220" target="_blank"><img src="products/index/index_4.gif" width="80" height="100" 
                  hspace=0 border=0></a></TD>
                        <TD align=right height=22> <TABLE cellSpacing=1 cellPadding=0 width="98%" bgColor=#f2f2f2 
                  border=0>
                            <TBODY>
                              <TR> 
                                <TD><FONT color=#000000>&nbsp;</FONT><FONT color=#ffffff><FONT 
                        color=#ffffff><IMG height=9 src="images/dot.gif" 
                        width=4></FONT></FONT><FONT color=#000000><IMG height=12 
                        src="images/pic111.gif" width=12> <B>商业高安全V-ASP </B></FONT></TD>
                                <TD width=34 bgColor=#f3f3f3><FONT color=#000000><A 
                        href="showsubs.asp?id=220" target="_blank"><FONT 
                        color=#666666><IMG height=20 src="images/B.GIF" 
                        width=32 border=0></FONT></A></FONT></TD>
                                <TD width=34 bgColor=#f3f3f3> <DIV align=center><FONT color=#000000><A 
                        href="http://dns99.cn/detail.php?gid=116&amp;nowcatid=6"></A><A 
                        href="add.asp?add=220"><FONT 
                        color=#666666><IMG height=20 src="images/M.GIF" 
                        width=32 
                  border=0></FONT></A></FONT></DIV></TD>
                              </TR>
                            </TBODY>
                          </TABLE></TD>
                      </TR>
                      <TR> 
                        <TD align=right height=22> <TABLE cellSpacing=1 cellPadding=0 width="98%" border=0>
                            <TBODY>
                              <TR> 
                                <TD width="55%" height=21><FONT color=#000000>价格：<FONT 
                        color=#999999><FONT color=#000000>568.00元/<FONT 
                        color=#000000><FONT color=#999999><FONT 
                        color=#000000><FONT color=#999999><FONT 
                        color=#000000>年</FONT></FONT></FONT></FONT></FONT></FONT></FONT></FONT></TD>
                                <TD width="45%" height=21><FONT color=#000000><FONT 
                        color=#999999><FONT color=#000000><FONT 
                        color=#999999><FONT color=#000000>详情</FONT></FONT>：</FONT><FONT 
                        color=#ff0000></FONT> <A 
                        href="showsubs.asp?id=220" target="_blank"><FONT 
                        color=#ff0000>点此查看</FONT></A> </FONT></FONT></TD>
                              </TR>
                            </TBODY>
                          </TABLE></TD>
                      </TR>
                      <TR> 
                        <TD align=right bgColor=#ffffff> <TABLE cellSpacing=2 cellPadding=0 width="98%" border=0>
                            <TBODY>
                              <TR> 
                                <TD><FONT color=#000000>送1个英文国际域名，20M Mssql数据库。2x300M空间任意分配网站空间和企业邮箱空间。支持ASP+ACCESS+FLASH...</FONT></TD>
                              </TR>
                            </TBODY>
                          </TABLE>
                          <FONT 
                  color=#000000><FONT color=#999999><FONT 
                  color=#000000></FONT></FONT></FONT></TD>
                      </TR>
                      <TR> 
                        <TD vAlign=top bgColor=#ffffff colSpan=2 height=4><FONT 
                  color=#ffffff><FONT color=#ffffff><IMG height=4 
                  src="images/dot.gif" width=8></FONT></FONT></TD>
                      </TR>
                    </TBODY>
                  </TABLE></td>
              </tr>
              <tr> 
                <td colspan="2" background="images/middle-2.gif"><img src="images/middle-2.gif" width="5" height="2"></td>
              </tr>
              <tr> 
                <td colspan="2"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20"><img src="images/gogo2.gif" width="20" height="30"></td>
                      <td><b><a href="China-Ecompany.asp">网站建设 特别推荐</a></b></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td colspan="2">超值之选，企业建站标准型，送1个英文国际域名250M X 2 空间，250M空间任意分配网站空间和企业邮箱空间( 
                  送 国际域名；免费 配置邮局，邮箱10M/个)。</td>
              </tr>
              <tr> 
                <td width="45%" valign="top">适用于想通过建立网站宣传产品或服务，提升企业形象，扩大品牌影响，拓展海内外潜在市场的企业。</td>
                <td width="55%" align="right"><a href="China-Ecompany.asp"><img src="images/web_8.gif" width="132" height="83" border="0"></a></td>
              </tr>
            </table></td>
          <td width="7">&nbsp;</td>
          <td width="208" valign="top" background="images/right-208x5.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="28" background="images/right-1.gif"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="14%" height="18" align="center" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="86%" valign="bottom"><font color="#FFFFFF"><b>最 
                        新 公 告</b></font></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><TABLE width=100% border=0 cellPadding=0 cellSpacing=0>
                    <TBODY>
                      <TR> 
                        <TD height="3" 
                vAlign=top 
                style="BORDER-right: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                      </TR>
                    </TBODY>
                  </TABLE></td>
              </tr>
              <tr> 
                <td align="right"><table width="207" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="2" valign="top" background="images/left-2-right.gif"><img src="images/left-2-right.gif" width="2" height="5"></td>
                      <td valign="top">
                        <!--#include file="bznew-index.asp" -->
                      </td>
                      <td width="7">&nbsp;</td>
                    </tr>
                  </table>
                  <TABLE width=100% border=0 cellPadding=0 cellSpacing=0>
                    <TBODY>
                      <TR> 
                        <TD height="3" 
                vAlign=top 
                style="BORDER-right: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                      </TR>
                    </TBODY>
                  </TABLE></td>
              </tr>
            </table>
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="208" height="28" background="images/right-1.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="14%" height="18" align="center" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="86%" valign="bottom"><font color="#FFFFFF"><b>业 
                        届 动 态</b></font></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><TABLE width=100% border=0 cellPadding=0 cellSpacing=0>
                    <TBODY>
                      <TR> 
                        <TD height="3" 
                vAlign=top 
                style="BORDER-right: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                      </TR>
                    </TBODY>
                  </TABLE></td>
              </tr>
              <tr> 
                <td height="150" align="right" valign="top"><table width="207" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="2" valign="top" background="images/left-2-right.gif"><img src="images/left-2-right.gif" width="2" height="5"></td>
                      <td>
					  <!--#include file="it-news-index.asp" -->
					  </td>
                      <td width="7">&nbsp;</td>
                    </tr>
                  </table>
                  <TABLE width=100% border=0 cellPadding=0 cellSpacing=0>
                    <TBODY>
                      <TR> 
                        <TD height="3" 
                vAlign=top 
                style="BORDER-right: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                      </TR>
                    </TBODY>
                  </TABLE> </td>
              </tr>
            </table>
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="208" height="28" background="images/right-1.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="14%" height="18" align="center" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="86%" valign="bottom"><font color="#FFFFFF"><b>产 
                        品 推 荐</b></font></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><TABLE width=100% border=0 cellPadding=0 cellSpacing=0>
                    <TBODY>
                      <TR> 
                        <TD height="3" 
                vAlign=top 
                style="BORDER-right: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                      </TR>
                    </TBODY>
                  </TABLE></td>
              </tr>
              <tr> 
                <td align="right" valign="top"><table width="207" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="2" valign="top" background="images/left-2-right.gif"><img src="images/left-2-right.gif" width="2" height="5"></td>
                      <td> 
                        <!--#include file="hosttj.asp"-->
                      </td>
                      <td width="7">&nbsp;</td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="208" height="244" valign="top" background="images/r_1.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td>&nbsp;</td>
                    </tr>
                    <tr> 
                      <td height="63" valign="top"><table width="100%" height="100%" border="0" cellpadding="6" cellspacing="1">
                          <tr> 
                            <td valign="bottom"><a href="#"><font color="#FF3300">[点 
                              此 进 入 演 示]</font></a></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td height="8"><img src="1.gif" width="1" height="1"></td>
                    </tr>
                    <tr> 
                      <td height="68"><table width="186" height="63" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td width="32%" height="68" rowspan="2">&nbsp;</td>
                            <td width="68%" height="36">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td height="23" align="center"><SELECT 
            onchange="MM_jumpMenu('parent',this,0)" name=select>
                                <option selected>-双IDC 加强型-</option>
                                <option>---</option>
                                <option value="showsubs.asp?id=401">双IDC加强型A</option>
                                <option value="showsubs.asp?id=402">双IDC加强型B</option>
                                <option value="showsubs.asp?id=403">双IDC加强型C</option>
                                <option value="showsubs.asp?id=404">双IDC加强型D</option>
                                <option value="showsubs.asp?id=405">双IDC加强型E</option>
                                <option value="showsubs.asp?id=406">双IDC加强型F</option>
                              </SELECT></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td height="8"><img src="1.gif" width="1" height="1"></td>
                    </tr>
                    <tr> 
                      <td height="66"><table width="186" height="60" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td width="32%" height="68" rowspan="2">&nbsp;</td>
                            <td width="67%" height="36">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td height="23" align="center"><SELECT 
            onchange="MM_jumpMenu('parent',this,0)" name=menu1>
                                <option selected>-名家为您推广-</option>
                                <option>---</option>
                                <option value="showsubs.asp?id=703">Google关键字</option>
                                <option value="showsubs.asp?id=702">百度竞价排名</option>
                                <option value="showsubs.asp?id=707">搜狐普通登录</option>
                                <option value="showsubs.asp?id=713">新浪快速登录</option>
                              </SELECT></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td height="8">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
  </tr>
</table>

<!--#include file="copyright.asp"-->
</body>
</html>