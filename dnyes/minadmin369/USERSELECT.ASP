<!--#include file="SHEI45FECSA.ASP" -->
<html>
<head>
<title>系统管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/southdns.css" type="text/css">
</head>
<body bgcolor="#FFFFCC" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" vspace="0" hspace="0">
  <tr bgcolor="#FFCC33"> 
    <td height="27">.:: 您可以在这里进行用户相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>




<form action="userselect.asp" method="post">

  <table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" width="87%" align="center">
    <tr bgcolor="#FFCC33"> 
      <td height="22" align="center">id</td>
      <td height="22" align="center">uid</td>
      <td align="center" bgcolor="#FFCC33">email</td>
      <td align="center">contactz</td>
      <td align="center">shengz</td>
      <td align="center">cityz</td>
      <td width="15%" align="center">userlevel</td>
      <td width="8%" align="center">prepay</td>
      <td width="16%" align="center">sumjifen</td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="22" bgcolor="#FFFF99" width="4%">
	  <input name="zid" id="zid" type="text" size="4"></td>
      <td height="22" bgcolor="#FFFF99" width="9%">
	  <input name="uid" id="uid" type="text" size="10"></td>
      <td width="16%" bgcolor="#FFFF99">
	  <input name="email" id="email" type="text" size="18"></td>
      <td width="10%">
	  <input name="contactz" id="contactz" type="text" size="12"> 
      </td>
      <td width="12%" bgcolor="#FFFF99">
	  <select name="shengz" id="shengz">
          <option value="" selected>请选择...</option>
          <option value="北京">北京</option>
          <option value="上海">上海</option>
          <option value="天津">天津</option>
          <option value="重庆">重庆</option>
          <option value="天津">天津</option>
          <option value="广东">广东</option>
          <option value="山东">山东</option>
          <option value="山西">山西</option>
          <option value="四川">四川</option>
          <option value="福建">福建</option>
          <option value="江苏">江苏</option>
          <option value="浙江">浙江</option>
          <option value="河北">河北</option>
          <option value="河南">河南</option>
          <option value="黑龙江">黑龙江</option>
          <option value="吉林">吉林</option>
          <option value="辽宁">辽宁</option>
          <option value="内蒙古">内蒙古</option>
          <option value="海南">海南</option>
          <option value="陕西">陕西</option>
          <option value="安徽">安徽</option>
          <option value="江西">江西</option>
          <option value="甘肃">甘肃</option>
          <option value="新疆">新疆</option>
          <option value="湖南">湖南</option>
          <option value="湖北">湖北</option>
          <option value="云南">云南</option>
          <option value="广西">广西</option>
          <option value="宁夏">宁夏</option>
          <option value="贵州">贵州</option>
          <option value="青海">青海</option>
          <option value="西藏">西藏</option>
          <option value="香港">香港</option>
          <option value="台湾">台湾</option>
          <option value="澳门">澳门</option>
        </select></td>
      <td width="10%">
	  <input name="cityz" type="text" id="cityz" size="12"></td>
      <td>
	  <select name="userlevel" id="userlevel">
          <option value="" selected>当前用户级别</option>
          <%
sqlt="select * from userlevel"
set rst=server.createobject("ADODB.Recordset")
rst.open sqlt,conn,1,1
do while not rst.eof
%>
          <option value="<%=rst("userlevel")%>"><%=rst("userlevel")%></option>
<%
rst.movenext
loop
rst.Close()
set rst=nothing
%>
        </select></td>
      <td>
	  <input name="prepay" id="prepay" type="text" size="6"></td>
      <td>
	  <input name="sumjifen" id="sumjifen" type="text" size="6"></td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="35" colspan="10"> <div align="center"> 
      <input name="selecting" type="submit" value="Select...">
        </div></td>
    </tr>
  </table>
    </form>

<%
zid=request("zid")
uid=request("uid")
email=request("email")
contactz=request("contactz")
shengz=request("shengz")
cityz=request("cityz")
userlevel=request("userlevel")
prepay=request("prepay")
sumjifen=request("sumjifen")

sql="select * from user where uid<>'' "

if trim(zid)<>"" then sql=sql + " and id>="&clng(zid)
if trim(uid)<>"" then sql=sql + " and uid='"&uid&"'"
if trim(email)<>"" then sql=sql + " and email='"&email&"'"
if trim(contactz)<>"" then sql=sql + " and contactz='"&contactz&"'"
if trim(shengz)<>"" then sql=sql + " and shengz='"&shengz&"'"
if trim(cityz)<>"" then sql=sql + " and cityz='"&cityz&"'"
if trim(userlevel)<>"" then sql=sql + " and userlevel='"&userlevel&"'"
if trim(prepay)<>"" then sql=sql + " and prepay>="&clng(prepay)
if trim(sumjifen)<>"" then sql=sql + " and sumjifen>="&clng(sumjifen)
%>

<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
%>
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="93%" align="center">
  <tr> 
    <td height="22" bgcolor="#FFCC33"> ID </td>
    <td bgcolor="#FFCC33">用户帐号</td>
    <td bgcolor="#FFCC33">联系人</td>
    <td bgcolor="#FFCC33">shengz</td>
    <td bgcolor="#FFCC33">cityz</td>
    <td bgcolor="#FFCC33">电话-1</td>
    <td bgcolor="#FFCC33">邮件地址</td>
    <td width="9%" bgcolor="#FFCC33">prepay</td>
    <td width="8%" bgcolor="#FFCC33">sumjifen</td>
    <td width="10%" bgcolor="#FFCC33">操作</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr> 
    <td width="7%" height="22" bgcolor="#FFFF99"><%=rs("id")%></td>
    <td width="10%" bgcolor="#FFFF99"><a href="userorderlog.asp?uid=<%=rs("uid")%>"><%=rs("uid")%></a></td>
    <td width="9%" bgcolor="#FFFF99"><%=rs("contactz")%></td>
    <td width="7%" bgcolor="#FFFF99"><%=rs("shengz")%></td>
    <td width="7%" bgcolor="#FFFF99"><%=rs("cityz")%></td>
    <td width="16%" bgcolor="#FFFF99"><%=rs("tel")%></td>
    <td width="8%" bgcolor="#FFFF99"><a href="mailto:<%=rs("email")%>"><%=rs("email")%></a></td>
    <td width="17%" bgcolor="#FFFF99"><%=rs("prepay")%></td>
    <td width="9%" bgcolor="#FFFF99"><%=rs("sumjifen")%></td>
    <td bgcolor="#FFFF99"><a href="edituser.asp?userid=<%=rs("id")%>">[查看|编辑]</a></td>
  </tr>
  <%
rs.movenext
loop
%>
</table>
<%
rs.close
set rs=nothing
%>
<!--#include file="copyright.asp" -->
</body>
</html>
