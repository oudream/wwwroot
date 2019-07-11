<%@ Language=VBScript %>
<%  
Option Explicit
Dim Sv
%>
<HTML>
<Head><title>显示服务器变量</title></Head>
<BODY> 
<table colspan=8 cellpadding=5 border=0>
  <tr> 
    <td align=CENTER bgcolor="#800000" width=20%> <font style="ARIAL NARROW" color="#ffffff" size="2">环境变量名</font></td>
    <td align=CENTER width=80% bgcolor="#800000"> <font style="ARIAL NARROW" color="#ffffff" size="2">结果</font></td>
  </tr>
<%
WITH Response
for each Sv In Request.ServerVariables
	.Write "<tr>" 
	.Write "<td bgcolor='f7efde' align=CENTER> <font style='ARIAL NARROW' size='2'>"
	.Write Sv
	.Write "</font></td>"
	.Write "<td bgcolor='f7efde' align=CENTER> <font style='ARIAL NARROW' size='2'>"
	.Write Request.ServerVariables(Sv)
	.Write "</font></td></tr>"
next
END WITH
%>
</table>
</BODY>
</HTML>
