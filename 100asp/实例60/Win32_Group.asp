<Html>
<Head>
<Title>��ʾ����Windows�û����ʺ�</Title>
</Head>
<Body>
<% 
  Dim objLocator, objService, objWEBMCol, objWEBM
  Set objLocator = CreateObject("WbemScripting.SWbemLocator") 
  Set objService = objLocator.ConnectServer
  '��ȡwin32_Group�������
  Set objWEBMCol = objService.InstancesOf("Win32_Group")
  Response.write "<H2>��ʾ����Windows�û����˺�:</H2><HR><UL>"
  '�г���
  For Each objWEBM in objWEBMCol
   Response.write "<LI>����: " & objWEBM.Name & _
   ", <BR>˵��: " & objWEBM.Caption & _
   ", <BR>����: " & objWEBM.Description & _
   ", <BR>����: " & objWEBM.Domain & _
   ", <BR>��ȫ��־: " & objWEBM.SID & _
   ", <BR>��ȫ��־����: " & objWEBM.SIDType & _
   ", <BR>״̬: " & objWEBM.Status & "<BR></LI>"
  Next
  Response.write "</UL>"
  Set objLocator = Nothing
  Set objService = Nothing
  Set objWEBMCol = Nothing
  Set objWEBM = Nothing
%>
</Body>
</html> 
