<html>
<title>
������
</title>
<body>
<form action=calculator.asp method=post>
������1: <input type=text name=num1><br>
������2: <input type=text name=num2><br>
<p>
ѡ����Ҫ���еĲ���<br>
<input type=radio name=operation value="��" checked>��<br>
<input type=radio name=operation value="��">��<br>
<input type=radio name=operation value="��">��<br>
<input type=radio name=operation value="��">��<br>
<input type=submit><input type=reset>
</form>
<hr>
<%
dim n1,n2,op
if request.form.count=0 then
	response.end 
end if
n1=request.form("num1")
n2=request.form("num2")
op=request.form("operation")
if op="��" then
    response.write n1
    response.write "+"
    response.write n2
    response.write "="
    response.write clng(n1)+clng(n2)
elseif op="��" then
    response.write n1
    response.write "-"
    response.write n2
    response.write "="
    response.write clng(n1)-clng(n2)
elseif op="��" then
    response.write n1
    response.write "*"
    response.write n2
    response.write "="
    response.write clng(n1)*clng(n2)
elseif op="��" then
    response.write n1
    response.write "/"
    response.write n2
    response.write "="
    response.write clng(n1)/clng(n2)
end if
%>
</body>
</html>