<html>
<head>
<title>图书管理</title>
</head>
<body bgcolor="#f0ffff" vlink=mediumslateblue>
<FORM action="index.asp" method=POST id=form1 name=form1>
<P align=left><FONT><FONT face=隶书><STRONG><FONT size=5><FONT 
size=4>图书查询</FONT>：</FONT></STRONG><BR>请输入作者：<BR>
<INPUT name=author style="HEIGHT: 22px; LEFT: 10px; TOP: 72px; WIDTH: 153px"><BR>
<INPUT id=submit1 name=submit1 type=submit value="查询图书"></FONT></FONT></P></FONT></FORM>
<FORM action="add.asp" id=FORM2 method=post name=FORM2><STRONG><FONT size=4><BR><FONT face=隶书>
添加新书：<BR></FONT></FONT></STRONG><FONT face=隶书>
书名：<BR><INPUT id=text1 name=text1><BR>
作者：<BR><INPUT id=text2 name=text2><BR>
<INPUT id=submit2 name=submit2 type=submit value=添加新书><BR></FONT></FORM>
<FORM action="del.asp" method=post id=form3 name=form3><BR><FONT><FONT face=隶书><FONT size=4><STRONG>删除图书：</STRONG>
<BR><INPUT id=text3 name=text3></FONT><BR>
<INPUT id=submit3 name=submit3 type=submit value=删除图书><BR></FONT></FONT></FORM>
<FORM action="lend.asp" method=post id=form4 name=form4><FONT 
size=4><STRONG><BR><FONT face=隶书>借阅图书：</FONT></STRONG></FONT><BR><FONT face=隶书>借阅人：<BR>
<INPUT id=text4 name=text4><BR>图书编号：<BR>
<INPUT id=text6 name=text6><BR>
<INPUT id=submit4 name=submit4 type=submit value=借阅图书><BR></FONT></FORM>
<FORM action="still.asp" method=post id=form5 name=form5><BR><FONT><FONT face=隶书><STRONG><FONT size=4>
归还图书：<BR><INPUT id=text5 name=text5></FONT></STRONG>&nbsp;<BR>
<INPUT id=submit5 name=submit5 type=submit value=借阅图书><BR></FONT></FONT>
</FORM>

</body></HTML>