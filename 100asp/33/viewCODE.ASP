<%
SUB PrintLine (ByVal strLine)
    strLine=server.HTMLEncode(strLine)
    strLine=replace(strLine,"&lt;%","<FONT COLOR=#ff0000>&lt;%")
    strLine=replace(strLine,"%&gt;","%&gt;</FONT>")
    strLine=replace(strLine,"&lt;SCRIPT","<FONT COLOR=#0000ff>&lt;SCRIPT",1,-1,1)
    strLine=replace(strLine,"&lt;/SCRIPT&gt;","&lt;/SCRIPT&gt;</FONT>",1,-1,1)
    strLine=replace(strLine,"&lt;!--","<FONT COLOR=#008000>&lt;!--",1,-1,1)
    strLine=replace(strLine,"--&gt;","--&gt;</FONT>",1,-1,1)
    Response.Write strLine
END SUB
Function ShowCode(filename)
    Dim strFilename
    Dim FileObject, oInStream, strOutput    
    strFilename = filename
    Set FileObject = Server.CreateObject("Scripting.FileSystemObject")
    Set oInStream = FileObject.OpenTextFile(strFilename, 1, 0, 0 )
    While NOT oInStream.AtEndOfStream
        strOutput = oInStream.ReadLine
        Call PrintLine(strOutput)
        Response.Write("<BR>")
    Wend
end function
%>
<HTML>
<HEAD>
<TITLE>ASP源码浏览器</TITLE>
</HEAD>
<BODY BGCOLOR=#FFFFFF>
<form action=viewcode.asp method=post>
请输入ASP文件名
<input type=file name=filename>
<input type=submit value="查看源程序">
</form>
<%
on error resume next
dim file
file=request.form("filename")
response.write file & "源程序如下<hr>"
if trim(file)<> "" then
  call showcode(file)
end if
%>
</FONT>
</BODY>
</HTML>
