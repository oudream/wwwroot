<%@ Language=VBScript %>
<HTML>
<HEAD>
<Title>��ȡ������ַ</Title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<h2>��ȡ������ַ</h2>
<%
set mAddress=server.CreateObject ("MyMacAdd.MacAddress")
Response.Write maddress.GetMACAddress 

%>

</BODY>
</HTML>
