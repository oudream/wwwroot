<%@ Language=VBScript %>
<HTML>
<TITLE>
订购水果
</TITLE>
<BODY>

请选择你所要订购的水果
<hr>
<FORM ACTION="order.asp">
<input name="fruit" type=checkbox value="苹果">苹果
<Br>
<input name="fruit" type=checkbox value="香蕉">香蕉
<Br>
<input name="fruit" type=checkbox value="菠萝">菠萝
<Br>
<input name="fruit" type=checkbox value="桔子">桔子
<input type=submit value="订购">
</FORM>
<hr>
<%
if Request.QueryString("fruit").Count=0 then
%>
你没有订购水果
<%else%>
你订购了
<% 
for each fruit in Request.QueryString("fruit")
	response.write "<br><font color=green>" & fruit & "</font>"
next
end if
%>
</BODY>
</HTML>