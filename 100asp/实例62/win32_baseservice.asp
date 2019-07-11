<html>
<head>
<title>NT服务查看器</title>
</head>
<body>
<%
   Dim objLocator, objService, objWEBMCol, objWEBM
   Set objLocator = CreateObject("WbemScripting.SWbemLocator") 
  Set objService = objLocator.ConnectServer
  Set objWEBMCol = objService.InstancesOf("Win32_BaseService")
  Response.write "<H2><center>NT服务查看器</center></H2><HR><UL>"
  For Each objWEBM in objWEBMCol
   Response.write "<LI>服务名称: " & objWEBM.Caption 
   response.write ", <BR>能被暂停: " & objWEBM.AcceptPause 
   response.write ", <BR>能被停止: " & objWEBM.AcceptStop 
   response.write ", <BR>服务描述: " & objWEBM.Description 
   response.write ", <BR>显示名称: " & objWEBM.DisplayName
   response.write ", <BR>文件路径: " & objWEBM.PathName 
   response.write ", <BR>服务类型: " & objWEBM.ServiceType 
   response.write ", <BR>已经启动: " & objWEBM.Started 
   response.write ", <BR>启动模式: " & objWEBM.StartMode 
   response.write ", <BR>当前状态: " & objWEBM.State & "<BR></LI>"
  Next
  Response.write "</UL>"
  Set objLocator = Nothing
  Set objService = Nothing
  Set objWEBMCol = Nothing
  Set objWEBM = Nothing
%>
</body>
</html>