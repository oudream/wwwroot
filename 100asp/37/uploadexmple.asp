<%@ Language=VBScript %>
<%Option Explicit%>
<html>
<title>
文件上载
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
    Response.Write "<center><b>祝贺你已经成功上载了" & Uploader.File.FileName & "文件</b><br></center><hr>"
    Response.Write "<table align=center border=0>"
    Response.Write "<tr><td><font color=red>文件名称:</font></td><td> " 
    Response.Write Uploader.File.FileName & "</td></tr>"
    Response.Write "<tr><td><font color=red>文件大小:</font></td><td> "
    Response.Write Uploader.File.FileSize & " bytes</td></tr>"
    Response.Write "<tr><td><font color=red>文件类型:</font></td><td> " 
    Response.Write Uploader.File.ContentType & "</td></tr>"
end if
%>
</body>
</html>