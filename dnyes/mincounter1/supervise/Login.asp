<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	if(Request.QueryString("logout").Count>0) logout();

	if(isPostBack){
		oConn.Open();
		doLogin();
		oConn.Close();
	}
	
	var bIsLogined = chkLogin(false);
	if(bIsLogined) doAlert("","Index.asp");
%>
<!--
	COCOON Counter 6
	Develop by Cocoon Studio. [ www.ccopus.com ]
	Product by Sunrise_Chen. (MSN:sunrise_chen@msn.com)
	
	Please don't remove this information, thanks.
-->
<html>
<head>
<title><%=WebTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="../style/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<!--#include file="inc_head.asp"-->
<form name="form1" method="post">
<table width="760" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
    <tr>
      <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td nowrap="nowrap" style="text-align:left">
				<span id="ItemTitle"><font face="webdings">8</font> 用户登录</span>
			</td>
            <td nowrap="nowrap" style="text-align:left">&nbsp;</td>
            <td width="180" align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><TABLE class=InnerTable cellSpacing=1 cellPadding=10 width="100%" 
        border=0>
        <TBODY>
          <TR>
            <TD  class=InnerHead align=middle rowSpan=3 width="175">
              <table cellspacing=0 cellpadding=3 
        border=0>
                <tbody>
                  <tr>
                    <td>Login ID </td>
                    <td style="PADDING-LEFT: 5pt">
                      <input type=text class=Input id=UID name=UID style="width:108px">
                    </td>
                  </tr>
                  <tr>
                    <td>Password</td>
                    <td style="PADDING-LEFT: 5pt">
                      <input type=password class=Input id=PWD name=PWD size="15" style="width:108px">
                    </td>
                  </tr>
                  <tr align="right">
                    <td  colspan=2>
                      <input class=GenButton type=submit value=登录 name="submit" style="width:38px">
&nbsp;
                <input class=GenButton type=reset value=取消 name="reset" style="width:38px">
                    </td>
                  </tr>
                </tbody>
            </table></TD>
            <TD class=InnerMain align=left colSpan=2 rowSpan=3 id="divLogin">
              <P>　　　　Welcome to use COCOON Counter 6 professional. please visit <a href="http://www.ccopus.com" target="_blank">www.ccopus.com</a> </P>
              <P>　　　　We will show you the place where dreams and life become one ... </P></TD>
          </TR>
          <TR> </TR>
          <TR> </TR>
        </TBODY>
      </TABLE></td>
    </tr>
    <tr>
      <td class="OuterFoot"></td>
    </tr>
  </table>
</form>
<!--#include file="../static/bott.htm"-->
</body>
</html>
