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
session("subs_id_9")=0
session("subs_id_10")=0
session("subs_id_11")=0
session("subs_id_12")=0
session("subs_id_13")=0
session("subs_id_14")=0
if request("subs_id_all")="no" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert('你还没有选择操作的商品,请先选择操作的商品. ');</SCRIPT>")
if request("subs_id_all")="yes" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert('成功添加');</SCRIPT>")
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="20">&nbsp;</td>
  </tr>
</table>
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="80" height="30"><p>名称</p></td>
          <td height="30">&nbsp;</td>
        </tr>
        <tr> 
          <td width="80" height="30"><p>CPU</p></td>
          <td height="30"> <iframe id=IFrame src="subs_send_in1.asp" frameborder=0 width=98% scrolling=no height=50></iframe> 
            &nbsp; </td>
        </tr>
        <tr> 
          <td width="80" height="30"><p>主板</p></td>
          <td height="30"> <iframe id=IFrame src="subs_send_in2.asp" frameborder=0 width=98% scrolling=no height=50></iframe> 
            &nbsp; </td>
        </tr>
        <tr> 
          <td width="80" height="30"><p>内存</p></td>
          <td height="30"> <iframe id=IFrame src="subs_send_in3.asp" frameborder=0 width=98% scrolling=no height=50></iframe> 
            &nbsp; </td>
        </tr>
        <tr> 
          <td width="80" height="30"><p>硬盘</p></td>
          <td height="30"> <iframe id=IFrame src="subs_send_in4.asp" frameborder=0 width=98% scrolling=no height=50></iframe> 
            &nbsp; </td>
        </tr>
        <tr> 
          <td width="80" height="30"><p>显示卡</p></td>
          <td height="30"> <iframe id=IFrame src="subs_send_in5.asp" frameborder=0 width=98% scrolling=no height=50></iframe> 
            &nbsp; </td>
        </tr>
        <tr> 
          <td width="80" height="30"><p>显示器</p></td>
          <td height="30"> <iframe id=IFrame src="subs_send_in6.asp" frameborder=0 width=98% scrolling=no height=50></iframe> 
            &nbsp; </td>
        </tr>
        <tr> 
          <td width="80" height="30"><p>光驱</p></td>
          <td height="30"> <iframe id=IFrame src="subs_send_in7.asp" frameborder=0 width=98% scrolling=no height=50></iframe> 
            &nbsp; </td>
        </tr>
        <tr> 
          <td width="80" height="30"><p>软驱</p></td>
          <td height="30"> <iframe id=IFrame src="subs_send_in8.asp" frameborder=0 width=98% scrolling=no height=50></iframe> 
            &nbsp; </td>
        </tr>
        <tr> 
          <td width="80" height="30"><p>键盘</p></td>
          <td height="30"> <iframe id=IFrame src="subs_send_in9.asp" frameborder=0 width=98% scrolling=no height=50></iframe> 
            &nbsp; </td>
        </tr>
        <tr> 
          <td width="80" height="30"><p>鼠标</p></td>
          <td height="30"> <iframe id=IFrame src="subs_send_in10.asp" frameborder=0 width=98% scrolling=no height=50></iframe> 
            &nbsp; </td>
        </tr>
        <tr> 
          <td width="80" height="30"><p>机箱电源</p></td>
          <td height="30"> <iframe id=IFrame src="subs_send_in11.asp" frameborder=0 width=98% scrolling=no height=50></iframe> 
            &nbsp; </td>
        </tr>
        <tr> 
          <td width="80" height="30"><p>音箱</p></td>
          <td height="30"> <iframe id=IFrame src="subs_send_in12.asp" frameborder=0 width=98% scrolling=no height=50></iframe> 
            &nbsp; </td>
        </tr>
        <tr> 
          <td colspan="2" align="center"><iframe id=IFrame src="subs_send_in13.asp" frameborder=0 width=98% scrolling=no height=50></iframe> 
            &nbsp; </td>
        </tr>
        <tr> 
          <td colspan="2" align="center"><iframe id=IFrame src="subs_send_in14.asp" frameborder=0 width=98% scrolling=no height=50></iframe> 
            &nbsp; </td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="100" align="right"><form name="form1" method="post" action="subs_send_submit.asp" target="_blank">
        <input type="submit" name="Submit" value='选定所配置的商品>>下一步'>
      </form>
      
    </td>
  </tr>
</table>
</body>
</html>
