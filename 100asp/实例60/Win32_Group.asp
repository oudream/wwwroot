<Html>
<Head>
<Title>显示所有Windows用户组帐号</Title>
</Head>
<Body>
<% 
  Dim objLocator, objService, objWEBMCol, objWEBM
  Set objLocator = CreateObject("WbemScripting.SWbemLocator") 
  Set objService = objLocator.ConnectServer
  '获取win32_Group服务组件
  Set objWEBMCol = objService.InstancesOf("Win32_Group")
  Response.write "<H2>显示所有Windows用户组账号:</H2><HR><UL>"
  '列出表单
  For Each objWEBM in objWEBMCol
   Response.write "<LI>名称: " & objWEBM.Name & _
   ", <BR>说明: " & objWEBM.Caption & _
   ", <BR>描述: " & objWEBM.Description & _
   ", <BR>域名: " & objWEBM.Domain & _
   ", <BR>安全标志: " & objWEBM.SID & _
   ", <BR>安全标志类型: " & objWEBM.SIDType & _
   ", <BR>状态: " & objWEBM.Status & "<BR></LI>"
  Next
  Response.write "</UL>"
  Set objLocator = Nothing
  Set objService = Nothing
  Set objWEBMCol = Nothing
  Set objWEBM = Nothing
%>
</Body>
</html> 
