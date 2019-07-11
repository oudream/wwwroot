<%@ LANGUAGE = VBScript %>
<% Option Explicit %>
<Html>
<title>猜数字游戏</title>
<body>
<%
      '设置页面不使用缓存
      Response.Expires = 0
%>
<%
dim GuessNum
on error resume next
GuessNum=Request("Number")
if GuessNum="" then GuessNum="0" End if
GuessNum=Clng(GuessNum)

Session("Count")=Session("Count") + 1
if Session("Count") < 10 and GuessNum <> session("Number") then

%>

<form action="guessNumber.asp">
    <input type="text" name="Number">
	<input type="submit" value="提交">
</form>
<hr>
<%
end if
if GuessNum<0 or guessNum>100 then
    Response.write "请输入1~100之间的整数"
elseif GuessNum=0 then
	session("Count") = 0
    Randomize
	session("Number") = Int(rnd * 100 + 1)
    Response.write "请输入1~100之间的整数"
elseif GuessNum > session("Number") then
	response.write "你猜的太大了"
elseif GuessNum < session("Number") then
	response.write "你猜的太小了"
elseif GuessNum = session("Number") then
	response.write "祝贺你,猜对了"
end if

Response.write "<br>共猜了" & Session("Count") & "次"
if Session("Count")=10 then
    Response.write "答案是" & Session("Number")
end if
%>
<a href="guessnumber.asp?Number=0">重新猜</a>
</body>
</html>
