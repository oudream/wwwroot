<html>
<head>
<title>�¼��鿴��</title>
</head>
<body>
<% 
  Dim objLocator, objService, objWEBMCol, objWEBM, vType
  Set objLocator = CreateObject("WbemScripting.SWbemLocator") 
 
  'Establish a connection to WMI

   Set objService = objLocator.ConnectServer

 
  'Get the Webm Service object
  Set objWEBMCol = objService.ExecQuery("SELECT * FROM " & _
 " win32_NTLogEvent WHERE LogFile='system'")
 
  Response.write "<H2><center>�¼��鿴��</center></H2><HR><UL>"
  'Enumerate 
%>
<table border=1 width=120%>
<tr>
  <td width=10%>�����</td><td width=8%>��Դ</td><td width=8%>�¼�</td><td width=10%>����</td><td>��Ϣ����</td><td width=20%>�û�</td>
</tr>
<%
  For Each objWEBM in objWEBMCol
 
 vType = ""
 if objWEBM.Type = "error"  then
    vType = "����" 
 elseif objWEBM.Type = "warning"  then
    vType = "����" 
 elseif objWEBM.Type = "information"  then
    vType = "��Ϣ" 
 else
    vType=objWEBM.Type
 end if
 
   Response.write "<tr><td>" & objWEBM.ComputerName & _
    "</td><td>" & objWEBM.SourceName & _
    "</td><td>" & objWEBM.EventCode & _
    "</td><td>" & vType & _
    "</td><td>" & left(objWEBM.Message,15) & _
    "����</td><td>" & objWEBM.User
  Next
  response.write "</td></tr></table>"
  Set objLocator = Nothing
  Set objService = Nothing
  Set objWEBMCol = Nothing
  Set objWEBM = Nothing

%>
</body>
</html>
 