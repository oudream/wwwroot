<HTML>
<BODY>
<TITLE>
用户登录
</TITLE>
<%
if Request.Form.Count=0 then
%>
请输入用户名和密码
<FORM ACTION="login.asp" METHOD="post">
    <Table border=0>
    <tr><td>用户名:</td>
        <td><INPUT TYPE=text NAME=username VALUE=""></td>
    </tr>
    <tr><td>密码:</td>
        <td><INPUT TYPE=password NAME=password  VALUE=""></td>
    </tr>
    </Table>
    <INPUT TYPE=Submit VALUE=确认提交>
    <INPUT TYPE=reset VALUE=重新输入>
</FORM>

<%else%>
<%
Dim user
dim pwd
user=Request.Form("username")
pwd=Request.Form("password")

if user="fenfang" then
    if pwd="1234" then
        Response.write "用户登录成功"
    else 
        Response.write "用户密码无效"
    end if
else
    Response.write "用户无效"
end if
end if
%>