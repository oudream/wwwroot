<%@ Language=VBScript %>
<HTML>
<HEAD>
<Title>访问INI文件组件</Title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<h2>访问INI文件组件</h2>
<%
set mIni=server.CreateObject ("MyIniFile.IniFile")
Response.Write  mini.GetIniFileString("d:\1.ini","company","phone")
%>

</BODY>
</HTML>
