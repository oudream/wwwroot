<%@ Language=VBScript %>
<HTML>
<TITLE>
����ˮ��
</TITLE>
<BODY>

��ѡ������Ҫ������ˮ��
<hr>
<FORM ACTION="order.asp">
<input name="fruit" type=checkbox value="ƻ��">ƻ��
<Br>
<input name="fruit" type=checkbox value="�㽶">�㽶
<Br>
<input name="fruit" type=checkbox value="����">����
<Br>
<input name="fruit" type=checkbox value="����">����
<input type=submit value="����">
</FORM>
<hr>
<%
if Request.QueryString("fruit").Count=0 then
%>
��û�ж���ˮ��
<%else%>
�㶩����
<% 
for each fruit in Request.QueryString("fruit")
	response.write "<br><font color=green>" & fruit & "</font>"
next
end if
%>
</BODY>
</HTML>