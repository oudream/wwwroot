<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<HTML>
<HEAD> <TITLE>文件浏览器</TITLE> </HEAD>
<BODY>
<TABLE width="100%" border=1 bordercolor="#000000" align="left" cellpadding="2" cellspacing="0">
<TR align="left" valign="top"  bgcolor="#800000" > 
<TD width="60%"><FONT color="#ffffff"><B><FONT size="2" face="宋体">文件名</FONT></B></FONT></TD>
<TD width="15%"><FONT color="#ffffff"><B><FONT size="2" face="宋体">大小</FONT></B></FONT></TD>
<TD width="25%"><FONT color="#ffffff"><B><FONT size="2" face="宋体">修改日期</FONT></B></FONT></TD>
</TR>
<%
Dim objFSO
Dim objFile
Dim objFolder
Dim sMapPath
Set objFSO = CreateObject("Scripting.FileSystemObject")
sMapPath ="c:\windows"
Set objFolder = objFSO.GetFolder(sMapPath)
For Each objFile In objFolder.Files
%> 
<TR align="left" valign="top" bordercolor="#999999" bgcolor='f7efde'> 
<TD> <FONT size="2" face="宋体" color="#000000"><A href="<% = sMapPath & "/" & objFile.Name %>"> 
<%
Response.Write objFile.Name
%>
</A>
</FONT>
</TD>
<TD>
<FONT size="2" face="宋体" color="#000000">
<%
If objFile.Size <1024 Then
Response.Write objFile.Size & " Bytes"
ElseIf objFile.Size < 1048576 Then
Response.Write Round(objFile.Size / 1024.1) & " KB"
Else
Response.Write Round((objFile.Size/1024)/1024.1) & " MB"
End If
%>
</FONT>
</TD>
<TD>
<FONT size="2" face="宋体" color="#000000">
<% 
Response.Write objFile.DateLastModified
%>
</FONT>
</TD>
</FONT>
</TD>
</TR>
<%
Next
%> 
</TABLE>
</BODY>
</HTML>