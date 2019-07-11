<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
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
<title>信网-数信科技 虚拟主机 建站空间 具有良好的速度、稳定性、可靠性、安全性，是企业信息发布和实现电子商务的网站平台</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="www.dnyes.com-host,虚拟主机,建站空间,具有良好的速度,稳定性,可靠性,安全性,是企业信息发布和实现电子商务的网站平台">
<meta name="description" content="www.dnyes.com-host,虚拟主机,建站空间,具有良好的速度,稳定性,可靠性,安全性,是企业信息发布和实现电子商务的网站平台">
</head>
 
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<!--#include file="top.asp"-->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
    <td><TABLE width=100% border=0 cellPadding=0 cellSpacing=0>
        <TBODY>
          <TR> 
            <TD 
                vAlign=top 
                style="BORDER-left: #000000 1px solid; BORDER-RIGHT: #000000 1px solid"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2"><img src="images/1-x.gif" width="754" height="56"></td>
                </tr>
                <tr> 
                  <td width="157" height="500" bgcolor="#f0f0f0">&nbsp;</td>
                  <td width="597" align="center"><img src="images/sitemap.gif" width="520" height="650" border="0" usemap="#Map"></td>
                </tr>
              </table></TD>
          </TR>
        </TBODY>
      </TABLE></td>
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
  </tr>
</table>
<!--#include file="copyright3.asp" -->
<map name="Map"><area shape="rect" coords="119,184,193,213" href="China-dedicatedhosting.asp">
  <area shape="rect" coords="119,221,193,252" href="China-ecompany.asp">
  <area shape="rect" coords="119,259,194,286" href="China-WebPromote.asp">
  <area shape="rect" coords="119,296,192,325" href="whois.asp">
  <area shape="rect" coords="119,332,193,361" href="emailtome.asp">
  <area shape="rect" coords="118,145,192,174" href="China-favorables.asp">
  <area shape="rect" coords="120,109,194,138" href="China-mailboxes.asp">
  <area shape="rect" coords="119,71,193,100" href="Virtual-Hosting-Service.asp">
  <area shape="rect" coords="120,35,194,64" href="domain.asp">
  <area shape="rect" coords="119,408,194,438" href="connect.asp">
  <area shape="rect" coords="120,447,192,474" href="banks.asp">
  <area shape="rect" coords="119,482,194,511" href="#">
  <area shape="rect" coords="117,520,194,548" href="#">
  <area shape="rect" coords="119,557,194,586" href="#">
  <area shape="rect" coords="225,483,323,510" href="index.asp">
  <area shape="rect" coords="368,66,472,96" href="dnyesfaq.asp">
  <area shape="rect" coords="367,103,472,132" href="mybase.asp">
  <area shape="rect" coords="368,141,472,169" href="#">
  <area shape="rect" coords="367,180,470,209" href="#">
  <area shape="rect" coords="368,215,470,244" href="#">
  <area shape="rect" coords="369,252,470,281" href="showallproduct.asp">
  <area shape="rect" coords="368,290,468,321" href="banks.asp">
  <area shape="rect" coords="368,328,471,356" href="emailtome.asp">
  <area shape="rect" coords="367,424,471,452" href="viewmyorders.asp">
  <area shape="rect" coords="368,463,470,491" href="domainmanage.asp">
  <area shape="rect" coords="369,497,472,527" href="usermodview.asp">
  <area shape="rect" coords="365,537,469,564" href="usermod.asp">
  <area shape="rect" coords="367,574,472,601" href="usermodpass.asp">
</map>
</body>
</html>