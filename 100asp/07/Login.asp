<HTML>
<BODY>
<TITLE>
�û���¼
</TITLE>
<%
if Request.Form.Count=0 then
%>
�������û���������
<FORM ACTION="login.asp" METHOD="post">
    <Table border=0>
    <tr><td>�û���:</td>
        <td><INPUT TYPE=text NAME=username VALUE=""></td>
    </tr>
    <tr><td>����:</td>
        <td><INPUT TYPE=password NAME=password  VALUE=""></td>
    </tr>
    </Table>
    <INPUT TYPE=Submit VALUE=ȷ���ύ>
    <INPUT TYPE=reset VALUE=��������>
</FORM>

<%else%>
<%
Dim user
dim pwd
user=Request.Form("username")
pwd=Request.Form("password")

if user="fenfang" then
    if pwd="1234" then
        Response.write "�û���¼�ɹ�"
    else 
        Response.write "�û�������Ч"
    end if
else
    Response.write "�û���Ч"
end if
end if
%>