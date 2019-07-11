<%@ Language=VBScript %>
<HTML>
<HEAD>
<Title>生成一个Guid</Title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<h2>生成一个Guid</h2>
<%
set mGuid=server.CreateObject ("myguid.mguid")
Response.Write  mguid.GetGUID 

%>

</BODY>
</HTML>
