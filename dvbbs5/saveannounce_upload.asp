<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<script>
if (top.location==self.location){
	top.location="index.asp"
}
</script>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Forum_css.asp"-->
</head>

<body <%=Forum_body(11)%>>
<form name="form" method="post" action="saveannouce_upfile.asp?boardid=<%=request("boardid")%>" enctype="multipart/form-data">
<table width="100%" border=0 cellspacing=0 cellpadding=0>
<tr><td class=tablebody1 valign=top height=40>
<%if Cint(GroupSetting(7))=0 then%>
��û���ڱ���̳�ϴ��ļ���Ȩ��
<%else%>
<input type="hidden" name="filepath" value="<%=Forum_Setting(64)%>">
<input type="hidden" name="act" value="upload">
ͼƬ
<input type="file" name="file1" size=10>
<input type="submit" name="Submit" value="�ϴ�" onclick="parent.document.forms[0].Submit.disabled=true,
parent.document.forms[0].Submit2.disabled=true;"> �����ϴ���
<select name="filetype" size=1>
<option value="">�ļ�����</option>
<%
dim iupload
iupload=split(Forum_upload,",")
for i=0 to ubound(iupload)
response.write "<option value="&iupload(i)&">"&iupload(i)&"</option>"
next
%>
</select>
  ���ƣ�<%=GroupSetting(40)%>����ÿ��<%=GroupSetting(44)%>K
<%end if%>
</td></tr>
</table>
</form>
</body>
</html>
<%
conn.close
set conn=nothing
%>
