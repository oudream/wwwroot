<%@ Language=VBScript %>
<%Option Explicit%>
<html>
<title>
�ļ�����
</title>
<body>
<!-- #include file="upload.asp" -->
<%
Dim Uploader, File
Set Uploader = New FileUploader
on error resume next
Uploader.Upload()
uploader.File.SaveToDisk "E:\"
if err.number=0 then
    Response.Write "<center><b>ף�����Ѿ��ɹ�������" & Uploader.File.FileName & "�ļ�</b><br></center><hr>"
    Response.Write "<table align=center border=0>"
    Response.Write "<tr><td><font color=red>�ļ�����:</font></td><td> " 
    Response.Write Uploader.File.FileName & "</td></tr>"
    Response.Write "<tr><td><font color=red>�ļ���С:</font></td><td> "
    Response.Write Uploader.File.FileSize & " bytes</td></tr>"
    Response.Write "<tr><td><font color=red>�ļ�����:</font></td><td> " 
    Response.Write Uploader.File.ContentType & "</td></tr>"
end if
%>
</body>
</html>