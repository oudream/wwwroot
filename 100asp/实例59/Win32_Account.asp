<html>
<head>
<title>NT�û��˺Ų鿴��</title>
</head>
<body>
<%
 
  Dim objLocator, objService, objWEBMCol, objWEBM
  Set objLocator = CreateObject("WbemScripting.SWbemLocator") 
  Set objService = objLocator.ConnectServer
  Set objWEBMCol = objService.InstancesOf("Win32_Account")
  Response.write "<H2><center>NT�û��˺Ų鿴��</center></H2><HR><UL>"
  For Each objWEBM in objWEBMCol
   Response.write "<LI>�˺�: " & objWEBM.Caption
   response.write "<BR>����: " & objWEBM.Description 
   response.write "<BR>����: " & objWEBM.Domain
   response.write "<BR>��ȫID: " & objWEBM.SID 
   response.write "<BR>��ȫID����: " & objWEBM.SIDType 
   response.write "<BR>״̬: " & objWEBM.Status & "<BR></LI>"
  Next
  Response.write "</UL>"
  Set objLocator = Nothing
  Set objService = Nothing
  Set objWEBMCol = Nothing
  Set objWEBM = Nothing

%>
</body>
</html>