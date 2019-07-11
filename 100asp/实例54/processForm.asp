<Html>
<Head>
<Title>
</Title>
</Head>
<Body>
<%
Function ConvertFormtoXML(strXMLFilePath, strFileName)
 Dim objDom
 Dim objRoot
 Dim objField
 Dim objFieldValue
 Dim objattID
 Dim objattTabOrder
 Dim objPI
 Dim x
 Set objDom = server.CreateObject("Microsoft.XMLDOM")
 objDom.preserveWhiteSpace = True
 Set objRoot = objDom.createElement("contact")
 objDom.appendChild objRoot
 For x = 1 To Request.Form.Count
  If instr(1,Request.Form.Key(x),"btn") = 0 Then
   Set objField = objDom.createElement("field")
   Set objattID = objDom.createAttribute("id")
   objattID.Text = Request.Form.Key(x)
   objField.setAttributeNode objattID
   Set objattTabOrder = objDom.createAttribute("taborder")
   objattTabOrder.Text = x
   objField.setAttributeNode objattTabOrder
   Set objFieldValue = objDom.createElement("field_value")
   objFieldValue.Text = Request.Form(x)
   objRoot.appendChild objField
   objField.appendChild objFieldValue
  End If
 Next 
 Set objPI = objDom.createProcessingInstruction("xml", "version='1.0'")
 objDom.insertBefore objPI, objDom.childNodes(0)
response.write "<Br>"
response.write strXMLFilePath
response.write strFileName
response.write "<hr>"
 objDom.save strXMLFilePath & "\" & strFileName
 Set objDom = Nothing
 Set objRoot = Nothing
 Set objField = Nothing
 Set objFieldValue = Nothing
 Set objattID = Nothing
 Set objattTabOrder = Nothing
 Set objPI = Nothing
End Function
On Error Resume Next
ConvertFormtoXML "c:","Contact.xml"
If err.number <> 0 then
 Response.write("在保存你提交的信息时发生错误.")
Else
 Response.write("你所提交的信息已经被保存.")
End If
%>
</body>
</html>