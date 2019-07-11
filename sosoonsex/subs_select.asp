<!--#include file="adminconn.asp" -->
<%
session("subs_id_1")=0
session("subs_id_2")=0
session("subs_id_3")=0
session("subs_id_4")=0
session("subs_id_5")=0
session("subs_id_6")=0
session("subs_id_7")=0
session("subs_id_8")=0
session("subs_pass_type")=request("passtype")
if session("subs_pass_type")="in" then passstr="进"
if session("subs_pass_type")="out" then passstr="出"
if request("subs_id_all")="no" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert('你还没有选择操作的商品,请先选择操作的商品. ');</SCRIPT>")
if request("subs_id_all")="yes" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert('成功添加');</SCRIPT>")
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="20">&nbsp;</td>
  </tr>
</table>
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
  <tr bgcolor="#FFFFFF"> 
    <td height="22" colspan="4" align="center" class="header">&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="50" colspan="4" class="header">请先选定所 <font color="#FF0000"><%=passstr%></font> 商品</td>
  </tr>
  <tr> 
    <td width="100" align="center" bgcolor="#FFFFFF">种类</td>
    <td width="100" align="center" bgcolor="#FFFFFF">品牌</td>
    <td width="150" align="center" bgcolor="#FFFFFF">型号</td>
    <td width="240" align="center" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4" align="center" bgcolor="#FFFFFF">
	<iframe id=IFrame src="subs_select_in1.asp" frameborder=0 width=98% scrolling=no height=50></iframe>&nbsp; 
	</td>
  </tr>
  <tr> 
    <td colspan="4" align="center" bgcolor="#FFFFFF">
	<iframe id=IFrame src="subs_select_in2.asp" frameborder=0 width=98% scrolling=no height=50></iframe>&nbsp; 
	</td>
  </tr>
  <tr> 
    <td colspan="4" align="center" bgcolor="#FFFFFF">
	<iframe id=IFrame src="subs_select_in3.asp" frameborder=0 width=98% scrolling=no height=50></iframe>&nbsp; 
	</td>
  </tr>
  <tr> 
    <td colspan="4" align="center" bgcolor="#FFFFFF">
	<iframe id=IFrame src="subs_select_in4.asp" frameborder=0 width=98% scrolling=no height=50></iframe>&nbsp; 
	</td>
  </tr>
  <tr> 
    <td colspan="4" align="center" bgcolor="#FFFFFF">
	<iframe id=IFrame src="subs_select_in5.asp" frameborder=0 width=98% scrolling=no height=50></iframe>&nbsp; 
	</td>
  </tr>
  <tr> 
    <td colspan="4" align="center" bgcolor="#FFFFFF">
	<iframe id=IFrame src="subs_select_in6.asp" frameborder=0 width=98% scrolling=no height=50></iframe>&nbsp; 
	</td>
  </tr>
  <tr> 
    <td colspan="4" align="center" bgcolor="#FFFFFF">
	<iframe id=IFrame src="subs_select_in7.asp" frameborder=0 width=98% scrolling=no height=50></iframe>&nbsp; 
	</td>
  </tr>
  <tr> 
    <td colspan="4" align="center" bgcolor="#FFFFFF">
	<iframe id=IFrame src="subs_select_in8.asp" frameborder=0 width=98% scrolling=no height=50></iframe>&nbsp; 
	</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="80" colspan="4" align="right"><input type="submit" name="Submit" value='选定所<%=passstr%>商品>>下一步' onClick="location='subs_pass.asp';">
      </td>
  </tr>
</table>
</body>
</html>
