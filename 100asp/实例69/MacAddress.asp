<%@ Language=VBScript %>
<HTML>
<HEAD>
<Title>获取网卡地址</Title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<h2>获取网卡地址</h2>
<%
set mAddress=server.CreateObject ("MyMacAdd.MacAddress")
Response.Write maddress.GetMACAddress 

%>

</BODY>
</HTML>
