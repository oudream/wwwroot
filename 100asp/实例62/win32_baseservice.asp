<html>
<head>
<title>NT����鿴��</title>
</head>
<body>
<%
   Dim objLocator, objService, objWEBMCol, objWEBM
   Set objLocator = CreateObject("WbemScripting.SWbemLocator") 
  Set objService = objLocator.ConnectServer
  Set objWEBMCol = objService.InstancesOf("Win32_BaseService")
  Response.write "<H2><center>NT����鿴��</center></H2><HR><UL>"
  For Each objWEBM in objWEBMCol
   Response.write "<LI>��������: " & objWEBM.Caption 
   response.write ", <BR>�ܱ���ͣ: " & objWEBM.AcceptPause 
   response.write ", <BR>�ܱ�ֹͣ: " & objWEBM.AcceptStop 
   response.write ", <BR>��������: " & objWEBM.Description 
   response.write ", <BR>��ʾ����: " & objWEBM.DisplayName
   response.write ", <BR>�ļ�·��: " & objWEBM.PathName 
   response.write ", <BR>��������: " & objWEBM.ServiceType 
   response.write ", <BR>�Ѿ�����: " & objWEBM.Started 
   response.write ", <BR>����ģʽ: " & objWEBM.StartMode 
   response.write ", <BR>��ǰ״̬: " & objWEBM.State & "<BR></LI>"
  Next
  Response.write "</UL>"
  Set objLocator = Nothing
  Set objService = Nothing
  Set objWEBMCol = Nothing
  Set objWEBM = Nothing
%>
</body>
</html>