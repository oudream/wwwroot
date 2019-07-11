<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<script language="JavaScript">
function CheckForm()
{
	if (document.addsubs.subsname.value.length == 0) {
		alert("请输入商品名.");
		document.addsubs.subsname.focus();
		return false;
	}

	if (document.addsubs.price.value.length == 0) {
		alert("请输入商品的价格.");
		document.addsubs.price.focus();
		return false;
	}if (document.addsubs.gys.value.length == 0) {
		alert("请输入供应商.");
		document.addsubs.gys.focus();
		return false;
	}if (document.addsubs.subsunit.value.length == 0) {
		alert("请输入商品单位.");
		document.addsubs.subsunit.focus();
		return false;
	}if (document.addsubs.photo.value.length == 0) {
		alert("请输入图片.");
		document.addsubs.photo.focus();
		return false;
	}if (document.addsubs.id.value.length == 0) {
		alert("请输入id.");
		document.addsubs.id.focus();
		return false;
	}
	return true;
}
</script>
<%

if request("action")="add" then
'开始添加商品
color=request("color")
id=request("id")
subs=request("subs")
subsname=request("subsname")
price=request("price")
area=request("area")
sampleurl=request("sampleurl")
subsunit=request("subsunit")
gys=request("gys")

if trim(color) ="" then color=" "

if trim(sampleurl) ="" then sampleurl=" "

photo=request("photo")
if trim(photo) ="" then photo=" "

other = replace(request("other"),chr(13)," ")
if other="" then other="此商品资料整理中"

brief=request("brief")
if brief="" then brief="此商品资料整理中"

'开始检测此产品名是否唯一
set rs=server.createobject("ADODB.Recordset")
sql="select * from subs where subsname='"&subsname&"' and gys='"&gys&"' order by idd desc"
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then
response.write"<SCRIPT language=JavaScript>alert('对不起，您所添加的商品其商品名不唯一,请核实！');"
response.write"javascript:history.go(-1)</SCRIPT>"
rs.close
set rs=nothing
response.end
end if
'结束检验，添加商品
conn.execute "insert into subs (id,subsname,price,other,brief,area,subsunit,gys,color,sampleurl,photo) values ("&clng(id)&",'"&subsname&"','"&price&"','"&other&"','"&brief&"','"&area&"','"&subsunit&"','"&gys&"','"&color&"','"&sampleurl&"','"&photo&"')"


'取得idd，添加BOOKidd,编号及推荐项目
set rs=server.createobject("ADODB.Recordset")
sql="select * from subs order by idd desc"
rs.open sql,conn,1,1
idd=rs("idd")
id=rs("id")
rs.close
set rs=nothing
subs=subs & "-" & id
conn.execute "update subs set bookbm='"&id&"',subs='"&subs&"' where idd=" & idd
if request("ifdisc")="y" then conn.execute "update subs set ifdisc=true where idd=" & idd
if request("iftj")="y" then conn.execute "update subs set iftj=true where idd=" & idd
'导入添加图片页面
response.write"<SCRIPT language=JavaScript>alert('商品资料已添加或修改成功！');"
response.write"javascript:history.go(-1)</SCRIPT>"
end if
%>
<%
area="网站建设"
anumber="0006"
subs=anumber
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
    <td height="27">.:: 您可以在这里进行商品相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>

<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="90%" align="center">
  <form action="addsubs-web.asp" method="post" name="addsubs" onSubmit="return CheckForm();">
    <tr bgcolor="#FFCC33"> 
      <td height="30" colspan="6"> 您现在操作的分类为：<a href="addsubs-step1.asp"><%=area%></a></td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="30" width="10%">产品全称：</td>
      <td width="34%" bgcolor="#FFFF99"> &nbsp; <input type="text" name="subsname" size="30" maxlength="45" class="form"> 
      </td>
      <td width="11%">产品单位：</td>
      <td width="8%"><input name="subsunit" type="text" class="form" value="年" size="3" maxlength="35"> 
      </td>
      <td width="12%" bgcolor="#FFFF99">供应商：</td>
      <td width="25%" bgcolor="#FFFF99"><select name="gys">
          <option value="信网 数信科技" selected>信网 数信科技</option>
          <%
sql="select * from gys"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
while not rs.eof
%>
          <option value="<%=rs("gys")%>"><%=rs("gys")%></option>
          <%
rs.movenext
wend
rs.Close()
%>
        </select></td>
    </tr>
    <tr bgcolor="#FFCC33"> 
      <td height="28">优惠价格：</td>
      <td> &nbsp; <input type="text" name="price" size="15" maxlength="45" class="form">
        RMB </td>
      <td>是否打折：</td>
      <td> <input type="checkbox" name="ifdisc" value="y"></td>
      <td bgcolor="#FFCC33">是否推荐：</td>
      <td bgcolor="#FFCC33">&nbsp; <input type="checkbox" name="iftj" value="y"></td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="31" valign="middle">颜色：</td>
      <td>&nbsp; <input name="color" type="text" class="form" size="10" maxlength="6"> 
      </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>演示：</td>
      <td><input name="sampleurl" type="text" class="form" size="30" maxlength="45" value=""></td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="35">图片：</td>
      <td height="30">&nbsp;
        <input name="photo" type="text" class="form" size="30" maxlength="45" value=""> 
      </td>
      <td height="30">ID：</td>
      <td height="30"><input name="id" type="text" class="form" size="10" maxlength="10"></td>
      <td height="30">&nbsp;</td>
      <td height="30">&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="35">简介：<br> </td>
      <td height="35" colspan="5"> <textarea name="brief" rows="10" cols="62"></textarea> 
      </td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="35" > <div align="center">详细说明：</div></td>
      <td height="35"colspan="5"> <textarea name="other" rows="10" cols="80"></textarea> 
      </td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="35" colspan="6"> <div align="center"> 
          <input type="submit" name="Submit" value="添加">
          &nbsp;&nbsp; 
          <input type="reset" name="Submit" value="重填">
          &nbsp;&nbsp; 
          <input type="button" name="Submit" value="后 退" onClick="history.go( -1 );return true;">
          <input type="hidden" name="action" value="add">
          <input type="hidden" name="area" value="<%=area%>">
          <input type="hidden" name="subs" value="<%=subs%>">
        </div></td>
    </tr>
  </form>
</table>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
