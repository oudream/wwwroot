<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>图片上传</title>
<LINK href=site.css rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0">
<form name="form" method="post" action="upfile.asp" enctype="multipart/form-data" >
	<br>
	<p align=center>
	<input type="hidden" name="CopyrightInfo" value="http://www.chinaasp.com">
	<input type="hidden" name="filepath" value="http://<%=xpurl%>/<%=FileUploadPath%>">
	<input type="hidden" name="act" value="upload">
	文件：<input type="file" name="file1" size=30>
	<input type="submit" name="Submit" value="上传" ><br> 类型：jpg,gif,png,bmp,swf，限制：1000K。<br><br>如果您在文档中插入了图片，请在添加文章页面的<font color=red>图片新闻</font>处选择“<font color=red>是</font>”。
</form>
</body>
</html>