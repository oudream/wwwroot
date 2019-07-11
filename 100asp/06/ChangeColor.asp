<%@ Language=VBScript %>
<html>
<title>
移动鼠标改变背景颜色
</title>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function MakeColor(ThisColor) {
document.bgColor = ThisColor;
}
//-->
</SCRIPT>
<center>
<table cellspacing=2 Border="0"> 
<tr>
<%
Dim I1, I2, I3 ' Looping variables for RGB Color
For I1 = 0 to 15 step 3
For I2 = 0 to 15 step 3
For I3 = 0 to 15 step 3
Color = Hex(I1) & Hex(I1) & Hex(I2) & Hex(I2) & Hex(I3) & Hex(I3)
%>
<td bgcolor="#<%=Color%>">
<a href="#" LANGUAGE=javascript OnMouseOver="return MakeColor('#<%=Color%>');">
<img src="clear.gif" width=10 height=10 border="0"></a>
</td>
<%
Next
Next
%>
</tr>
<tr>
<%
Next
%>
</tr>
</table>
</center>
</html> 

