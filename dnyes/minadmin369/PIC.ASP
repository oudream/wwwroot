<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
area=request("area")
subsname=request("subsname")
subs=request("subs")
id=request("id")
%>
<script language="JavaScript">
    function See(hh,ff) {
        hh.src=ff.value
}
</Script>
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

<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="65%" align="center">
  <Form name="upload_file" enctype="multipart/form-data" action="picok.asp" method=post>
    <tr> 
      <td height="50" colspan="2" bgcolor="#FFCC33">您现在操作的分类为：<a href="addsubs-step3.asp?sarea=<%=sarea%>&bigarea=<%=bigarea%>"><%=area%></a><br>
        此产品的名称为：<%=subsname%>,产品编号为：<%=subs%> </td>
    </tr>
    <tr> 
      <td height="20" bgcolor="#FFFF99"> <input type=hidden name="id" value="<%=id%>"> 
        <input type=hidden name="FileUploadStart"> &nbsp; <input type="hidden" name="file_up1" size="30"> 
        &nbsp; <input type="file" name="file_up2" size="30"> <input type=hidden name="FileUploadEnd"> 
      </td>
      <td width="30%" align="center" valign="middle" bgcolor="#FFFF99">&nbsp; <input type="hidden" name="b1" value="预览1（大）" onClick="See(this.form.picview,this.form.file_up1)" class="form"> 
        &nbsp; <input type="button" name="b2" value="预览3（小）" onClick="See(this.form.picview,this.form.file_up2)" class="form">
      </td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="38" colspan="2"> <div align="center"><br>
          <img src="../image/image.gif" name="picview" border="0"><br>
          <br>
        </div></td>
    </tr>
    <tr> 
      <td height="35" colspan="2" bgcolor="#FFFFCC"> <div align="center"> 
          <input type="submit" name="Submit" value="添加">
          &nbsp;&nbsp; 
          <input type="reset" name="Submit" value="重填">
        </div></td>
    </tr>
  </form>
</table>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
