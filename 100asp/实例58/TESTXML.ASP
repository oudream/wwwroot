<html>
<head>
<title>XMLС����</title>
</head>
<body>
<h2><center>XMLС����</center></h2>
<%
set xml = Server.CreateObject("Microsoft.XMLDOM")
xml.async = false
xml.load(server.mappath("testXML.xml"))
set xsl = Server.CreateObject("Microsoft.XMLDOM")
xsl.async = false
xsl.load(server.mappath("testXSL.xsl"))
Response.Write(xml.transformNode(xsl))
%>
</body>
</html>