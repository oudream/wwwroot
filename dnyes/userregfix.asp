<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
<!--#include file="mdf.asp" -->
<html>
<head>
<title>用户确定</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
dim zadduser, zuid, zpwd, zcom_cn, zcom_en, zcontact_cn, zcontact_en, zu_or_i, ztechcc, ztechsp, ztechsp_cn, ztechen, ztechen_cn, zcityz, zcitye, zaddr, zaddr_en, zpostnumber, ztel1_gov_code, ztel1_area_code, ztel1, ztel1_ext_code,  ztel2_gov_code, ztel2_area_code, ztel2, ztel2_ext_code,  zfax_gov_code, zfax_area_code, zfax, zfax_ext_code, zqq, zemail, zmemo
zuid=trim(request("uid"))
zpwd=request("pwd")
zcom_cn=request("com_cn")
zcom_en=request("com_en")
zcontact_cn=request("contact_cn")
zcontact_en=request("contact_en")
zu_or_i=request("u_or_i")
ztechcc=request("techcc")
ztechsp=request("techsp")
ztechsp_cn=request("techsp_cn")
ztechen=request("techen")
ztechen_cn=request("techen_cn")
zcityz=request("cityz")
zcitye=request("citye")
zaddr=request("addr")
zaddr_en=request("addr_en")
zpostnumber=request("postnumber")
ztel1_gov_code=request("tel1_gov_code")
ztel1_area_code=request("tel1_area_code")
ztel1=request("tel1")
ztel1_ext_code=request("tel1_ext_code")
ztel2_gov_code=request("tel2_gov_code")
ztel2_area_code=request("tel2_area_code")
ztel2=request("tel2")
ztel2_ext_code=request("tel2_ext_code")
zfax_gov_code=request("fax_gov_code")
zfax_area_code=request("fax_area_code")
zfax=request("fax")
zfax_ext_code=request("fax_ext_code")
zqq=request("qq")
zemail=request("email")
zmemo=request("memo")
zusertype=request("usertype")
zpwdb=md5(zpwd)
%>
<%
if zu_or_i =0 then 
sexstr="女"
else
sexstr="男"
end if

if trim(ztechsp)<>"" then
provcnstr=ztechsp
else
provcnstr=ztechsp_cn
end if

if trim(ztechen)<>"" then
provenstr=ztechen
else
provenstr=ztechen_cn
end if

if trim(zusertype)<>"" then
	if zusertype="buss" then
	zusertypestr="企业"
	else
	zusertypestr="单位"
	end if
end if   

if trim(ztel1_ext_code)<>"" then
tel1str=cstr(ztel1_gov_code)&"-"&cstr(ztel1_area_code)&"-"&cstr(ztel1)&"-"&cstr(ztel1_ext_code)
else 
tel1str=cstr(ztel1_gov_code)&"-"&cstr(ztel1_area_code)&"-"&cstr(ztel1)
end if

if trim(ztel2_ext_code)<>"" then
tel2str=cstr(ztel2_gov_code)&"-"&cstr(ztel2_area_code)&"-"&cstr(ztel2)&"-"&cstr(ztel2_ext_code)
else
tel2str=cstr(ztel2_gov_code)&"-"&cstr(ztel2_area_code)&"-"&cstr(ztel2)
end if

if trim(zfax_ext_code)<>"" then
faxstr=cstr(zfax_gov_code)&"-"&cstr(zfax_area_code)&"-"&cstr(zfax)&"-"&cstr(zfax_ext_code)
else
faxstr=cstr(zfax_gov_code)&"-"&cstr(zfax_area_code)&"-"&cstr(zfax)
end if
%>
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
                                    <td valign="bottom" background="images/host-2.gif">用 
                                      户 注 册 >> 资 料 确 认</td>
                                  </tr>
                                </table>
                                <table width="75%" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td height="6"><img src="1.gif" width="1" height="1"></td>
                                  </tr>
                                </table>
                                <table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                                  <tr> 
                                    <td bgcolor="#f5f5f5"> 提示：<br> <br>
                                      1.注册为我们的会员用户后您可以在线订购我们的产品和服务<br> <br>
                                      2.你可以对于您的订单做在线的管理<br> <br>
                                      3.当您忘记密码的时候这些资料可以帮您顺利的找回密码<br> <br>
                                      所以：<font color="#FF0033">请您认真填写下列资料</font>. 
                                      （带<b><font color="#FF3333">*</font></b>为必须填写的项目） 
                                    </td>
                                  </tr>
                                </table>
                                <br>
                                <table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                                  <tr> 
                                    <td width="112" bgcolor="#f5f5f5">会员帐号：</td>
                                    <td width="363" bgcolor="#FFFFFF"><%=zuid%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">会员密码：</td>
                                    <td bgcolor="#FFFFFF"><%=zpwd%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5"><%=zusertypestr%>名称（中文）：</td>
                                    <td bgcolor="#FFFFFF"><%=zcom_cn%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5"><%=zusertypestr%>名称（英文）：</td>
                                    <td bgcolor="#FFFFFF"><%=zcom_en%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">联系人（中文）：</td>
                                    <td bgcolor="#FFFFFF"><%=zcontact_cn%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">联系人（英文）：</td>
                                    <td bgcolor="#FFFFFF"><%=zcontact_en%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">联系人性别：</td>
                                    <td bgcolor="#FFFFFF"><%=sexstr%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">国家或地区：</td>
                                    <td bgcolor="#FFFFFF"><%=ztechcc%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">省份(中文)：</td>
                                    <td bgcolor="#FFFFFF"><%=provcnstr%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">省份(英文)：</td>
                                    <td bgcolor="#FFFFFF"><%=provenstr%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">城市（中文）：</td>
                                    <td bgcolor="#FFFFFF"><%=zcityz%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">城市（英文）： </td>
                                    <td bgcolor="#FFFFFF"><%=zcitye%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">通信地址（中文）：</td>
                                    <td bgcolor="#FFFFFF"><%=zaddr%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">通信地址（英文）：</td>
                                    <td bgcolor="#FFFFFF"><%=zaddr_en%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">邮政编码：</td>
                                    <td bgcolor="#FFFFFF"><%=zpostnumber%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">联系电话1：</td>
                                    <td bgcolor="#FFFFFF"><%=tel1str%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">联系电话2：</td>
                                    <td bgcolor="#FFFFFF"><%=tel2str%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">传真：</td>
                                    <td bgcolor="#FFFFFF"><%=faxstr%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">OICQ/ICQ/MSN：</td>
                                    <td bgcolor="#FFFFFF"><%=zqq%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">电子邮件：</td>
                                    <td bgcolor="#FFFFFF"><%=zemail%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">备　注：</td>
                                    <td bgcolor="#FFFFFF"><%=zmemo%></td>
                                  </tr><tr bgcolor="#FFFFFF">
                                    <td colspan="2">&nbsp;</td>
                                  </tr>
                                  <tr bgcolor="#f5f5f5"> 
                                    <td height="40" colspan="2" align="center"> <br>
                                      <form name="ADDUser" method="POST" action="userregsucc.asp" >
                                        <input type="hidden" name="uid"  value="<%=zuid%>">
                                        <input type="hidden" name="pwd"  value="<%=zpwd%>">
                                        <input type="hidden" name="pwdb"  value="<%=zpwdb%>">
                                        <input name="com_cn" type="hidden"   value="<%=zcom_cn%>">
                                        <input name="com_en" type="hidden"   value="<%=zcom_en%>">
                                        <input type="hidden" name="contact_cn"  value="<%=zcontact_cn%>">
                                        <input name="contact_en" type="hidden"   value="<%=zcontact_en%>">
                                        <input name="u_or_i" type="hidden" value="<%=sexstr%>" >
                                        <input name="TECHCC" type="hidden" value="<%=ztechcc%>" >
                                        <input name="TECHSP" type="hidden" value="<%=ztechsp%>">
                                        <input name="TECHSP_CN" type="hidden" value="<%=ztechsp_cn%>" >
                                        <input name="TECHEN" type="hidden" value="<%=ztechen%>">
                                        <input name="TECHEN_CN" type="hidden" value="<%=ztechen_cn%>">
                                        <input type="hidden" name="cityz" value="<%=zcityz%>">
                                        <input type="hidden" name="citye" value="<%=zcitye%>">
                                        <input type="hidden" name="addr" value="<%=zaddr%>">
                                        <input type="hidden" name="addr_en" value="<%=zaddr_en%>">
                                        <input type="hidden" name="postnumber" value="<%=zpostnumber%>">
                                        <input type="hidden" name="tel1"   value="<%=tel1str%>">
                                        <input type="hidden" name="tel2"   value="<%=tel2str%>">
                                        <input type="hidden" name="fax"  value="<%=faxstr%>">
                                        <input type="hidden" name="QQ"  value="<%=zqq%>">
                                        <input type="hidden" name="email"  value="<%=zemail%>">
                                        <input type="hidden" name="memo"  value="<%=zmemo%>">
                                        <input class="botton" type="submit" name="ok" value="提交申请">&nbsp;&nbsp;&nbsp;&nbsp;
                                        <INPUT name="button" TYPE="button" class="botton" OnClick="history.go( -1 );return true;" VALUE="返回修改">
                                        <input type="hidden" name="adduser" value="true">
                                        <input type="hidden" name="usertype" value="<%=zusertype%>">
                                      </form></td>
                                  </tr>
                                </table> </td>
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
<!--#include file="copyright.asp"-->
</body>
</html>
