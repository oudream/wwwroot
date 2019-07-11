<%
Response.buffer=false
Server.ScriptTimeOut=150
%>
<Html>
<Head>
<Title>ÂúÌìĞÇ</Title>
</Head>
<Body bgcolor=#000000>
<%
dim startx
dim k,i
dim nextsecond
for k=1 to 60
    nextsecond=dateadd("s",2,time)
    do while time<nextsecond
    loop
    Randomize
    star=int(50*rnd())
    for i=1 to star
    Response.Write "<font color=000000>&nbsp;</font>"
    Next
    Response.Write "<font color=#00ffff>*</font>"
    if k mod 3=2 then response.write "<br>"
next
%>
</body>
</html>

