<%@ Language=VBScript %>
<html>
<head>
<title>实例 多层目录显示</title>
</head>
<body><font size=+2><STRONG>多层目录显示</STRONG></font><P>
中国行政区
<%
dim conn
dim rstemp1,rstemp2
dim sqltemp1,sqltemp2
conn = "DBQ=" + server.mappath("genus.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set rstmep1 = server.createobject("adodb.recordset")  
set rstmep2 = server.createobject("adodb.recordset")  
sqltemp1="select * from genus"
rstmep1.open sqltemp1, conn, 1, 1
%>  
<%do while not rstmep1.eof%>  
<DIV ID=Tree<%=rstmep1("id")%> CLASS="Folder">  
<img ID="Tree<%=rstmep1("id")%>Fldr" src="picture/close.bmp"> <%=rstmep1("name")%>
</DIV>
<DIV ID=Tree<%=rstmep1("id")%>cont STYLE="display:none">
<%  
set rstmep2 = server.createobject("adodb.recordset")  
sqltemp2="select * from book where genus_id=" & rstmep1("id")  
rstmep2.open sqltemp2, conn, 1, 1
%><table><%
do while not rstmep2.eof  
%>
<tr><td>&nbsp;&nbsp;<img src="picture/text.bmp"></td><td><%=rstmep2("name")%></td></tr>
<%
rstmep2.movenext  
loop
%></table>
</DIV>
<%rstmep1.movenext
 loop
 rstmep1.close
 rstmep2.close
 %>  
<SCRIPT LANGUAGE="VBScript">
Sub document_onClick()
If window.event.srcElement.className = "Folder" Then
strNode = window.event.srcElement.ID + "Cont"
strFldImg = window.event.srcElement.ID + "Fldr" 
If document.All(strNode).style.display = "" Then
  document.All(strNode).style.display = "none"
  document.All(strFldImg).Src = "picture\close.bmp"
Else
  document.All(strNode).style.display = ""
  document.All(strFldImg).Src = "picture\open.bmp"
End If
End If
End Sub
</SCRIPT> 
</td>  
</tr>
</table>
<%
set rs1 = nothing
set rs = nothing
 %>
</body>
</html>