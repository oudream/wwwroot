<html>
<head>
<title>
文件权限管理
</title>
</head>
<body>
<% Set pmck = Server.CreateObject("MSWC.PermissionChecker") %>
测试物理路径的访问权限<br>
c:\boot.ini:<%= pmck.HasAccess("c:\boot.ini") %>
<br>
测试虚拟路径的访问权限<br>
/test/1.txt:<%= pmck.HasAccess("/test/1.txt") %>
</body>
</html>
