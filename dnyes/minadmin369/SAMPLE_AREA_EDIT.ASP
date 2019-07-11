<script language="JavaScript">

function checkform()
{
	if (form_administrator.sample_area_name.value.length == 0) {
		alert("请输入虚拟空间名称");
		form_administrator.sample_area_name.focus();
		return false;
		}
	if (form_administrator.addr.value.length == 0) {
		alert("请输入虚拟空间地址");
		form_administrator.addr.focus();
		return false;
		}
	if (form_administrator.addr.value.length == 0) {
		alert("请输入虚拟空间图片");
		form_administrator.addr.focus();
		return false;
		}
	if (form_administrator.memo.value.length == 0) {
		alert("请输入虚拟空间说明");
		form_administrator.memo.focus();
		return false;
		}
	return true;
}

</script>

<!--#include file="SHEI45FECSA.ASP" -->
<%
'添加公告
if request("action")="edit" then

sample_area_name=request("sample_area_name")
sql="select * from sample_area where sample_area_name='" &sample_area_name&"' and id<>"& request("id")
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
if not (rs.bof or rs.eof) then 
response.Write("<table width=100% height=100% border=0 align=center cellpadding=0 cellspacing=0><tr><td align=center>此日期已经存在，请选择编辑<br><br></td></tr></table>")
response.End()
rs.close
set rs=nothing
end if

addr=request("addr")
img=request("img")
memo=request("memo")

conn.execute "update sample_area set sample_area_name='"&sample_area_name&"',addr='"&addr&"',img='"&img&"',memo='"&memo&"' where id=" & request("id")

response.Write("<SCRIPT LANGUAGE=JavaScript> alert('编辑成功');</SCRIPT>")
end if 
%>
<%
sql="select * from sample_area where id=" & request("id")
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
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
    <td height="27">.:: 您可以在这里进行DELEDDOMAIN相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="65%" align="center">
  <form action="sample_area_EDIT.ASP" method="post" name="form_administrator" id="form_administrator"  onSubmit="return checkform();">
    <tr bgcolor="#FFCC33"> 
      <td height="22" width="16%">虚拟空间名称</td>
      <td><input maxlength=100 name="sample_area_name" id="sample_area_name" size=50 class="form" value="<%=rs("sample_area_name")%>"></td>
    </tr>
    <tr bgcolor="#FFCC33"> 
      <td width="16%" height="22">虚拟空间地址：</td>
      <td>http://
        <input maxlength=100 name="addr" id="addr" size=50 class="form" value="<%=rs("addr")%>"> 
      </td>
    </tr>
    <tr bgcolor="#FFCC33"> 
      <td width="16%" height="22">图片：</td>
      <td><input maxlength=100 name="img" id="img" size=50 class="form" value="<%=rs("img")%>"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="45">说明：</td>
      <td height="45"><textarea name="memo" cols="60" rows="8" class="form" id="memo"><%=rs("memo")%>
</textarea></td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="45" colspan="2" align="center"> <input type="submit" name="Submit" value="确 定"> 
        &nbsp;&nbsp; <input type="reset" name="Submit" value="取 消"> &nbsp;&nbsp; 
        <input type="button" name="Submit" value="后 退" onClick="history.go( -1 );return true;"> 
        <input type="hidden" name="id" id="id" value="<%=request("id")%>">
        <input type="hidden" name="action" value="edit"> </td>
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
