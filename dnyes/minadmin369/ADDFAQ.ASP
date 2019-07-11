<!--#include file="SHEI45FECSA.ASP" -->
<html>
<head>
<title>系统管理中心</title>
<%
'添加公告
if request("action")="add" then

subject=request("subject")
mytype=request("mytype")
newstype=request("newstype")
subject=request("subject")
color=request("color")
if color="" then color=" "
message=request("message")
message=replace(message,"'","’")
conn.execute "insert into myfaq (subject,message,idate,newstype,mytype,color) values ('"&subject&"','"&message&"','"&date()&"','"&newstype&"','"&mytype&"','"&color&"')"

'取得ID，添加BOOKID,编号及推荐项目
set rs=server.createobject("ADODB.Recordset")
sql="select * from myfaq order by id desc"
rs.open sql,conn,1,1
id=rs("id")
rs.close
set rs=nothing

if request("iftj")="y" then conn.execute "update myfaq set ifshow=true where id=" & id
if request("iftj")="" then conn.execute "update myfaq set ifshow=false where id=" & id
response.Redirect("addfaqok.asp")
end if 
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/southdns.css" type="text/css">
</head>
<body bgcolor="#FFFFCC" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" vspace="0" hspace="0">
  <tr bgcolor="#FFCC33"> 
    <td height="27">.:: 您可以在这里进行FAQ/BASE相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="90%" align="center">
  <form action="ADDFAQ.ASP" method="post" name="addnews">
    <tr bgcolor="#FFCC33"> 
      <td height="22" width="19%">&nbsp;FAQ/BASE标题：</td>
      <td colspan="3">
<input maxlength=45 name="subject" size=40 class="form"> 
      </td>
    </tr>
    <tr bgcolor="#FFCC33"> 
      <td height="22" width="19%">FAQ/BASE类型：</td>
      <td width="28%">
<select name="newstype">
          <option value="域名">域名 </option>
          <option value="虚拟主机">虚拟主机 </option>
          <option value="企业邮箱">企业邮箱 </option>
          <option value="网站建设">网站建设 </option>
          <option value="套餐">套餐 </option>
          <option value="租用托管">租用托管 </option>
          <option value="网站推广">网站推广 </option>
          <option value="域名知识" selected>域名知识 </option>
          <option value="虚拟主机知识">虚拟主机知识 </option>
          <option value="企业邮箱知识">企业邮箱知识 </option>
          <option value="网站建设知识">网站建设知识 </option>
          <option value="套餐知识">套餐知识 </option>
          <option value="租用托管知识">租用托管知识 </option>
          <option value="网站推广知识">网站推广知识 </option>
        </select> </td>
      <td width="20%">FAQ/BASE颜色：</td>
      <td width="33%"><input  name="color" size="12" maxlength=6 class="form"></td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="22"> FAQ OR BASE： </td>
      <td colspan="3">
		<select name="mytype">
          <option value="faq">FAQ </option>
          <option value="base" selected>BASE </option>
        </select>
      </td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="45"> <div align="center">是否推荐：</div></td>
      <td height="45" colspan="3">
        <input type="checkbox" name="iftj" value="y"></td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="45">&nbsp;FAQ/BASE内容：</td>
      <td height="45" colspan="3"><textarea name="message" cols="90" class="form" rows="20">
</textarea></td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="45" colspan="4"> <input type="submit" name="Submit" value="确 定"> 
        &nbsp;&nbsp; <input type="reset" name="Submit" value="取 消"> &nbsp;&nbsp; 
        <input type="button" name="Submit" value="后 退" onClick="history.go( -1 );return true;"> 
        <input type="hidden" name="action" value="add"> </td>
    </tr>
  </form>
</table>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
