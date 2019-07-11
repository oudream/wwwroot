<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
'修改公告
if request("action")="modi" then

subject=request("subject")
mytype=request("mytype")
newstype=request("newstype")
message=request("message")
color=request("color")
if color="" then color=" "
conn.execute "update myfaq set subject='"&subject&"',message='"&message&"',newstype='"&newstype&"',mytype='"&mytype&"',color='"&color&"' where id=" & request("id")
if request("iftj")="y" then conn.execute "update myfaq set ifshow=true where id=" & request("id")
if request("iftj")="" then conn.execute "update myfaq set ifshow=false where id=" & request("id")
response.redirect "viewfaq.asp"
end if 
%>
<html>
<head>
<title>系统管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/southdns.css" type="text/css">
</head>
<body bgcolor="#FFFFCC" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" vspace="0" hspace="0">
  <tr bgcolor="#FFCC33"> 
    <td height="27">.:: 您可以在这里进行站内公告相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<%
sql="select * from myfaq where id=" & request("id")
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%> 
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="90%" align="center">
  <form action="EDITFAQ.ASP" method="post" name="news">
    <tr bgcolor="#FFCC33"> 
      <td height="33" width="14%">&nbsp;FAQ标题：</td>
      <td colspan="3"> &nbsp; <input maxlength=45 name=subject size=60 class="form" value="<%=rs("subject")%>"> -->>ID:<%=rs("id")%>
      </td>
    </tr>
    <tr bgcolor="#FFCC33"> 
      <td height="36" width="14%">FAQ OR BASE：</td>
      <td width="35%"> 
        <select name="mytype">
          <option value="<%=rs("mytype")%>" selected><%=rs("mytype")%> </option>
          <option value="faq">FAQ </option>
          <option value="base">BASE </option>
        </select></td>
      <td width="15%">COLOR:</td>
      <td width="36%"><input maxlength=6 name=color id=color size=10 class="form" value="<%=rs("color")%>"></td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="36">FAQ类型：</td>
      <td colspan="3"><select name="newstype">
          <option value="<%=rs("newstype")%>" selected><%=rs("newstype")%> </option>
          <option value="域名">域名 </option>
          <option value="虚拟主机">虚拟主机 </option>
          <option value="企业邮箱">企业邮箱 </option>
          <option value="网站建设">网站建设 </option>
          <option value="套餐">套餐 </option>
          <option value="租用托管">租用托管 </option>
          <option value="网站推广">网站推广 </option>
          <option value="域名知识">域名知识 </option>
          <option value="虚拟主机知识">虚拟主机知识 </option>
          <option value="企业邮箱知识">企业邮箱知识 </option>
          <option value="网站建设知识">网站建设知识 </option>
          <option value="套餐知识">套餐知识 </option>
          <option value="租用托管知识">租用托管知识 </option>
          <option value="网站推广知识">网站推广知识 </option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="34">是否推荐：</td>
      <td height="34" colspan="3"><%if rs("ifshow") then tiftj="checked" %><input type="checkbox" name="iftj" value="y" <%=tiftj%>></td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="45">&nbsp;FAQ内容：</td>
      <td height="45" colspan="3"><textarea name="message" cols="90" class="form" rows="20">
<%=rs("message")%>
</textarea></td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="45" colspan="4"><input type="submit" name="Submit" value="确 定"> 
        &nbsp;&nbsp; <input type="reset" name="Submit" value="取 消"> &nbsp;&nbsp; 
        <input type="button" name="Submit" value="后 退" onClick="history.go( -1 );return true;"> 
        <input type="hidden" name="action" value="modi"> <input type="hidden" name="ID" value="<%=request("id")%>"> 
      </td>
    </tr>
  </form>
</table>
<%
rs.close
set rs=nothing
%> 
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
