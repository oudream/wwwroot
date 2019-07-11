<html>
<body bgcolor=#c1f7d8>
<br><br>
<p>
<dl>
  <dd><div align=center style="font-size:xx-large"><font color="crimson">
  <p>BBS论坛</p></font></div></dd>
</dl>
<br><br>
<p align=center style="font-size:xx-large"><font size=4>欢迎访问BBS论坛</font></p>
<center>
你是第
<%
dim visitors
whichfile=server.mappath("counter/counter.txt")
set fs=server.createobject("Scripting.FileSystemObject")
set thisfile=fs.opentextfile(whichfile)
visitors=thisfile.readline
thisfile.close
countlen=len(visitors)
for i=1 to 5-countlen
    response.write "<img src=counter/0.gif>"
next
for i=1 to countlen
    response.write "<img src=counter/" & mid(visitors,i,1) & ".gif></img>"
next
visitors=visitors+1
set out=fs.createtextfile(whichfile)
out.writeline(visitors)
out.close
set fs=nothing
%>
位访问者
</center>
<hr>
<table align=center border=0 width=630>
   <tr>
      <td width=20% align=center><a href="register.htm">[会员注册]</a></td>
      <td width=20% align=center><a href="login.htm">[会员登录]</a></td>
      <td width=20% align=center><a href="visitor.htm">[游客访问]</a></td>
      <td width=20% align=center><a href="help.htm">[帮助信息]</a></td>
      <td width=20% align=center><a href="leavemessage.htm">[请你留言]</a></td>
   </tr>
</table>
</body>
</html>