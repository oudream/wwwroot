<%@ LANGUAGE = VBScript %>
<% Option Explicit %>
<Html>
<title>��������Ϸ</title>
<body>
<%
      '����ҳ�治ʹ�û���
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
	<input type="submit" value="�ύ">
</form>
<hr>
<%
end if
if GuessNum<0 or guessNum>100 then
    Response.write "������1~100֮�������"
elseif GuessNum=0 then
	session("Count") = 0
    Randomize
	session("Number") = Int(rnd * 100 + 1)
    Response.write "������1~100֮�������"
elseif GuessNum > session("Number") then
	response.write "��µ�̫����"
elseif GuessNum < session("Number") then
	response.write "��µ�̫С��"
elseif GuessNum = session("Number") then
	response.write "ף����,�¶���"
end if

Response.write "<br>������" & Session("Count") & "��"
if Session("Count")=10 then
    Response.write "����" & Session("Number")
end if
%>
<a href="guessnumber.asp?Number=0">���²�</a>
</body>
</html>
