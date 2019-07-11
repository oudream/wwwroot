<html>
<title>
计算器
</title>
<body>
<form action=calculator.asp method=post>
操作数1: <input type=text name=num1><br>
操作数2: <input type=text name=num2><br>
<p>
选择你要进行的操作<br>
<input type=radio name=operation value="加" checked>加<br>
<input type=radio name=operation value="减">减<br>
<input type=radio name=operation value="乘">乘<br>
<input type=radio name=operation value="除">除<br>
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
if op="加" then
    response.write n1
    response.write "+"
    response.write n2
    response.write "="
    response.write clng(n1)+clng(n2)
elseif op="减" then
    response.write n1
    response.write "-"
    response.write n2
    response.write "="
    response.write clng(n1)-clng(n2)
elseif op="乘" then
    response.write n1
    response.write "*"
    response.write n2
    response.write "="
    response.write clng(n1)*clng(n2)
elseif op="除" then
    response.write n1
    response.write "/"
    response.write n2
    response.write "="
    response.write clng(n1)/clng(n2)
end if
%>
</body>
</html>