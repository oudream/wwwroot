<%@ Language=VBScript %>
<HTML>
<HEAD>
<Title>注册表访问组件</Title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<h2>注册表访问组件</h2>
<%
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
set mRegistKey=server.CreateObject ("MyAccessRegist.Regist")
mRegistKey.UpdateKey  HKEY_CURRENT_USER,"TEST","TEST","TEST"
Response.Write mregistkey.GetKeyValue( HKEY_CURRENT_USER,"TEST","TEST")
%>

</BODY>
</HTML>
