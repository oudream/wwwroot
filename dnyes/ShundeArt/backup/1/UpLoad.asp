<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ͼƬ�ϴ�</title>
<LINK href=site.css rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0">
<form name="form" method="post" action="upfile.asp" enctype="multipart/form-data" >
	<br>
	<p align=center>
	<input type="hidden" name="CopyrightInfo" value="http://www.chinaasp.com">
	<input type="hidden" name="filepath" value="http://<%=xpurl%>/<%=FileUploadPath%>">
	<input type="hidden" name="act" value="upload">
	�ļ���<input type="file" name="file1" size=30>
	<input type="submit" name="Submit" value="�ϴ�" ><br> ���ͣ�jpg,gif,png,bmp,swf�����ƣ�1000K��<br><br>��������ĵ��в�����ͼƬ�������������ҳ���<font color=red>ͼƬ����</font>��ѡ��<font color=red>��</font>����
</form>
</body>
</html>