<!--#include file="upload_hizi.inc" -->
<!--#include file=config.asp -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ϴ�ͼƬ</title>
<style type="text/css">
<!--
a {  font-family: "����"; font-size: 9pt; font-style: normal; line-height: 13pt; font-weight: normal; font-variant: normal; text-transform: none; color: <%=fontcolor%>; text-decoration: none}
a:hover {  font-family: "����"; font-size: 9pt; font-style: normal; line-height: 13pt; font-weight: normal; font-variant: normal; text-transform: none; color: <%=fontcolor%>; text-decoration: underline}
td {  font-family: "����"; font-size: 9pt; font-style: normal; line-height: 13pt; font-weight: normal; font-variant: normal; text-transform: none; color: <%=fontcolor%>}
br {  font-family: "����"; font-size: 9pt; font-style: normal; line-height: 13pt; font-weight: normal; font-variant: normal; text-transform: none; color: <%=fontcolor%>}
.bk { font-size: 9pt; border: 1px <%=xcolor%> solid}
body {  font-family: "����"; font-size: 9pt; font-style: normal; line-height: 13pt; font-weight: normal; font-variant: normal; text-transform: none}
.an {  font-family: "����"; font-size: 9pt; background-color: <%=bgcolor%>; border: 1px <%=xcolor%> solid; color: <%=fontcolor%>}
.xzy {  border: <%=xcolor%> solid; border-width: 0px 1px 1px}
.zx {  border: <%=xcolor%> solid; border-width: 0px 0px 1px 1px}
.sxz {  border: <%=xcolor%> solid; border-width: 1px 0px 1px 1px}
.s {  border: <%=xcolor%>; border-style: solid; border-top-width: 1px; border-right-width: 0px; border-bottom-width: 0px; border-left-width: 0px}
.y {  border: <%=xcolor%>; border-style: solid; border-top-width: 0px; border-right-width: 1px; border-bottom-width: 0px; border-left-width: 0px}
.font {  font-family: "Arial Black"; font-size: 14pt; color: <%=fontcolor%>}
.x {  border: <%=xcolor%>; border-style: solid; border-top-width: 0px; border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px}
.z {  border: <%=xcolor%>; border-style: solid; border-top-width: 0px; border-right-width: 0px; border-bottom-width: 0px; border-left-width: 1px}
.sx {  border: <%=xcolor%>; border-style: solid; border-top-width: 1px; border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px}
-->
</style>
<body leftmargin="0" topmargin="0">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td align="center">
<script language="Javascript">
function minipic(smileface)
{
window.opener.document.form.gif_url.value=smileface;
}
</script>
<%
set upload=new upload_hizi
set file=upload.file("file1")
formPath="upimg/"

fileExt=lcase(right(file.filename,3))
if fileExt<>"swf" and fileExt<>"gif" then
%>
<script language="VBScript" type="text/VBScript">
msgbox "���ͼƬֻ�����ϴ�Gif��Swf(Flash)�������͵��ļ���������ѡ���ϴ���"
 history.back(1)
              </script>
<%
response.end
end if
%>
<%
if file.filesize>1024*1024*2 then
%>
<script language="VBScript" type="text/VBScript">
msgbox "���ͼƬ��С������2M���ڣ�������ѡ���ϴ���"
 history.back(1)
              </script>
<%
end if
randomize
ranNum=int(90000*rnd)+10000
filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
if file.FileSize>0 then
file.SaveAs Server.mappath(FileName)
end if
response.write "һ��СΪ <font color=red>"&file.filesize&"</font> �ֽڵ� <font color=red>"&fileExt&"</font> ͼƬ�ϴ��ɹ���<br><br>A.�������� [<a href=Javascript:minipic('"&DqUrl&"/"&filename&"');>�Զ�����ͼƬURL��ͼƬ��ַ��</a>]<br><br>B.�ֶ�����ͼƬURL��ַ��ͼƬ��ַ��,URL����:<br><input size=30 value="&DqUrl&"/"&filename&"><br><br>[<a href='"&DqUrl&"/"&filename&"'>���Ԥ��</a>]"
%>
</td>
</tr>
</table>
</body>