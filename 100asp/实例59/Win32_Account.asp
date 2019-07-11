<html>
<head>
<title>NT用户账号查看器</title>
</head>
<body>
<%
 
  Dim objLocator, objService, objWEBMCol, objWEBM
  Set objLocator = CreateObject("WbemScripting.SWbemLocator") 
  Set objService = objLocator.ConnectServer
  Set objWEBMCol = objService.InstancesOf("Win32_Account")
  Response.write "<H2><center>NT用户账号查看器</center></H2><HR><UL>"
  For Each objWEBM in objWEBMCol
   Response.write "<LI>账号: " & objWEBM.Caption
   response.write "<BR>描述: " & objWEBM.Description 
   response.write "<BR>域名: " & objWEBM.Domain
   response.write "<BR>安全ID: " & objWEBM.SID 
   response.write "<BR>安全ID类型: " & objWEBM.SIDType 
   response.write "<BR>状态: " & objWEBM.Status & "<BR></LI>"
  Next
  Response.write "</UL>"
  Set objLocator = Nothing
  Set objService = Nothing
  Set objWEBMCol = Nothing
  Set objWEBM = Nothing

%>
</body>
</html>