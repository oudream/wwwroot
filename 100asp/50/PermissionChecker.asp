<html>
<head>
<title>
�ļ�Ȩ�޹���
</title>
</head>
<body>
<% Set pmck = Server.CreateObject("MSWC.PermissionChecker") %>
��������·���ķ���Ȩ��<br>
c:\boot.ini:<%= pmck.HasAccess("c:\boot.ini") %>
<br>
��������·���ķ���Ȩ��<br>
/test/1.txt:<%= pmck.HasAccess("/test/1.txt") %>
</body>
</html>
