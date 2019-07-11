<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<script language="JavaScript">
function CheckForm()
{
	if (document.modisubs.subsname.value.length == 0) {
		alert("请输入商品名.");
		document.modisubs.subsname.focus();
		return false;
	}

	if (document.modisubs.price.value.length == 0) {
		alert("请输入商品的价格.");
		document.modisubs.price.focus();
		return false;
	}if (document.modisubs.gys.value.length == 0) {
		alert("请输入供应商.");
		document.modisubs.gys.focus();
		return false;
	}if (document.modisubs.subsunit.value.length == 0) {
		alert("请输入商品单位.");
		document.modisubs.subsunit.focus();
		return false;
	}if (document.modisubs.photo.value.length == 0) {
		alert("请输入图片.");
		document.modisubs.photo.focus();
		return false;
	}if (document.modisubs.id.value.length == 0) {
		alert("请输入id.");
		document.modisubs.id.focus();
		return false;
	}
	return true;
}
</script>
<%
if request("action")="modi" then
'开始修改资料
idd=request("idd")
id=request("id")
subsname=request("subsname")
price=request("price")
subsnamey=request("subsnamey")
subsunit=request("subsunit")
gys=request("gys")
gysy=request("gysy")
color=request("color")
if color="" then color=" "
sampleurl=request("sampleurl")
if trim(sampleurl) ="" then sampleurl=" "

photo=request("photo")
if trim(photo) ="" then photo=" "

other=request("other")
if other="" then other="此商品无详细说明"

brief=request("brief")
if brief="" then brief="此商品无详细说明"

'以下为备用资料
area=request("area")

'开始检测此产品名是否唯一
'注，由于存在不修改商品名的可能性，所以先行进行判断subsname和subsnamey是否相等
if subsname<>subsnamey or gys<>gysy then
set rs=server.createobject("ADODB.Recordset")
sql="select * from subs where subsname='"&subsname&"' and gys='"&gys&"' order by idd desc"
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then
response.write"<SCRIPT language=JavaScript>alert('对不起，您修改后的商品其商品名不唯一,请核实！');"
response.write"javascript:history.go(-1)</SCRIPT>"
rs.close
set rs=nothing
response.end
end if
end if
'结束检验，修改
conn.execute "update subs set subsname='"&subsname&"',price='"&price&"',id="&clng(id)&",gys='"&gys&"',other='"&other&"',photo='"&photo&"',brief='"&brief&"',subsunit='"&subsunit&"',sampleurl='"&sampleurl&"' where idd=" & idd 
if request("ifdisc")="y" then conn.execute "update subs set ifdisc=true where idd=" & idd
if request("ifdisc")="" then conn.execute "update subs set ifdisc=false where idd=" & idd
if request("iftj")="y" then conn.execute "update subs set iftj=true where idd=" & idd
if request("iftj")="" then conn.execute "update subs set iftj=false where idd=" & idd
if request("ifdeal")="y" then conn.execute "update subs set ifdeal=true where idd=" & idd
if request("ifdeal")="" then conn.execute "update subs set ifdeal=false where idd=" & idd
'返回页面
response.write"<SCRIPT language=JavaScript>alert('商品资料已添加或修改成功！');"
response.write"javascript:history.go(-2)</SCRIPT>"
url="editsubs.asp?idd=" & idd
end if
%>
<%
idd=request("idd")
set rs=server.createobject("ADODB.Recordset")
sql="select * from subs where idd=" & idd
rs.open sql,conn,1,1
%>

<html>
<head>
<title>商城系统管理中心</title>
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
  <form action="editsubs.asp" method="post" name="modisubs" onSubmit="return CheckForm();">
    <tr bgcolor="#FFCC33"> 
      <td height="50" colspan="6"> 您现在操作的分类为：<%=rs("area")%><br> <a href="showsubsimage.asp?idd=<%=idd%>" target="_blank">查看此商品图片资料</a>&nbsp;&nbsp; 
        <a href="pic.asp?idd=<%=idd%>&area=<%=rs("area")%>&subs=<%=rs("subs")%>&subsname=<%=rs("subsname")%>">修改此商品的图片资料 
        </a></td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="30" width="10%">产品全称：</td>
      <td bgcolor="#FFFF99" width="35%">&nbsp; <input type="text" name="subsname" size="30" maxlength="45" class="form" value="<%=rs("subsname")%>"> 
      </td>
      <td width="10%">产品单位：</td>
      <td width="10%" align="center">&nbsp; <input type="text" name="subsunit" size="3" maxlength="35" class="form" value="<%=rs("subsunit")%>"> 
      </td>
      <td width="10%" bgcolor="#FFFF99">供应商：</td>
      <td width="25%" bgcolor="#FFFF99"><input type="text" name="gys" size="15" maxlength="45" class="form" value="<%=rs("gys")%>"></td>
    </tr>
    <tr bgcolor="#FFCC33"> 
      <td height="30">优惠价格：</td>
      <td bgcolor="#FFCC33">&nbsp; <input type="text" name="price" size="15" maxlength="45" class="form" value="<%=rs("price")%>">
        RMB </td>
      <td>是否对代理打折：</td>
      <td>&nbsp; <%
ifdisc=""
if rs("ifdisc")=true then ifdisc="checked"
%> <input type="checkbox" name="ifdisc" value="y" <%=ifdisc%>> 
      </td>
      <td bgcolor="#FFCC33">是否推荐：</td>
      <td bgcolor="#FFCC33">&nbsp; <%
iftj=""
if rs("iftj")=true then iftj="checked"
%> <input type="checkbox" name="iftj" value="y" <%=iftj%>></td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="30" valign="middle">颜色：</td>
      <td bgcolor="#FFFF99">&nbsp; <input type="text" name="color" size="10" maxlength="6" class="form" value="<%=rs("color")%>"> 
      </td>
      <td bgcolor="#FFFF99">是否在处理中：</td>
      <td bgcolor="#FFFF99"> <%
ifdeal=""
if rs("ifdeal")=true then ifdisc="checked"
%> <input type="checkbox" name="ifdeal" value="y" <%=ifdeal%>></td>
      <td bgcolor="#FFFF99">演示：</td>
      <td bgcolor="#FFFF99"><input name="sampleurl" type="text" class="form" idd="sampleurl" size="30" maxlength="45" value="<%=rs("sampleurl")%>"></td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="31" bgcolor="#FFCC33">图片：</td>
      <td height="31" bgcolor="#FFCC33">&nbsp;
        <input name="photo" type="text" class="form" size="30" maxlength="45" value="<%=rs("photo")%>"></td>
      <td height="31" bgcolor="#FFCC33">ID：</td>
      <td height="31" bgcolor="#FFCC33"><input name="id" type="text" class="form" value="<%=rs("id")%>" size="10" maxlength="10"></td>
      <td height="31" bgcolor="#FFCC33">&nbsp;</td>
      <td height="31" bgcolor="#FFCC33">&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="64"> <div align="center">简介： </div></td>
      <td height="64" colspan="5"><textarea name="brief" rows="10" cols="62"><%=rs("brief")%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="64"> <div align="center">详细说明：</div></td>
      <td height="64" colspan="5"><textarea name="other" rows="10" cols="80"><%=rs("other")%></textarea></td>
    </tr>
    <tr align="center" bgcolor="#FFFFCC"> 
      <td height="64" colspan="6"> <input type="submit" name="Submit" value="修改"> 
        &nbsp;&nbsp; <input type="reset" name="Submit" value="重填"> &nbsp;&nbsp; 
        <input type="button" name="Submit" value="后 退" onClick="history.go( -1 );return true;"> 
        <input type="hidden" name="action" value="modi"> <input type="hidden" name="idd" value="<%=rs("idd")%>"> 
        <input type="hidden" name="subsnamey" value="<%=rs("subsname")%>"> 
        <input type="hidden" name="area" value="<%=rs("area")%>"> 
        <input type="hidden" name="gysy" value="<%=rs("gys")%>"> 
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
