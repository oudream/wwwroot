<head><title>网友评论：</title></head>
<body bgcolor="#f0ffff">
<table border="1" width="751" height="83">
<tr><td align=middle><font size=-1>（以下网友留言只代表网友个人观点，不代表本网站的观点）</font></td></tr>
<%
dim rs
dim sql
dim selected
dim conn
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="select * from message order by id desc"
rs.open sql,conn,1,1
if rs.EOF and  rs.BOF then
	%></TD>
  <TR><td></td></TR><%
	Response.Write "请您发表评论......"
else
	do while not rs.eof
%>
<TR>
<td >&nbsp;<IMG src="hand.gif">网友&nbsp;<%=rs("user")%>       于<%=rs("publishtime")%>发表评论</td></TR>
	<tr><td>&nbsp;<%=rs("content")%></td></tr><%
	rs.movenext
	loop
end if
set rs=nothing
set conn=nothing
%></TD></TR>
</table>
<div align="center">
  <center>&nbsp;
  </center>
<CENTER>
  <table border="0" cellpadding="4" cellspacing="4" width="462" height="201">
  <form action="savemessage.asp"></CENTER>  
  <TBODY>
    <tr>
	<td width="431" colspan="4">
      <P align=center>说明：本留言版只有会员才能提交发言，如果你<BR>还没有成为会员，请先<A href="user_register.asp" target=_blank><font size="2" face="仿宋_GB2312" color="#0000ff"><b>注册</b></font></A></P></td></tr>
    <tr>
      <td width="431" colspan="4"><p align="left">请您发表的评论</p>
       </td>
    </tr>
  <center>
   <tr> 
      <td width="452" height="7">您的昵称：<INPUT id=textname name=textname style="HEIGHT: 20px; WIDTH: 113px"> 
      您的密码：<INPUT id=textpassword name=textpassword style="HEIGHT: 20px; WIDTH: 96px"  
type=password></td> 
    </tr> 
    <tr> 
      <td><TEXTAREA id=textcontent name=textcontent rows=1 style="HEIGHT: 150px; WIDTH: 447px"></TEXTAREA></td> 
    </tr> 
    <tr> 
      <td width="452" height="14"> 
      <P  
      align=center><input type="submit" value="确定" name="action">
</P></td> 
    </tr></FORM></TBODY>
    <tr><td colspan="4" width="431">
    <P>请您注意：
    <BR>1,&nbsp;尊重网上道德，遵守中华人民共和国的各项有关法律法规  
	<BR>2,&nbsp;承担一切因您的行为而直接或间接导致的民事或刑事法律  
	<BR>3,&nbsp;本站管理人员有权保留或删除其管辖留言中的任意内容  
	<BR>4,&nbsp;您在留言板发表的作品，本网站有权在网站内转载可引用  
	<BR>5,&nbsp;参与本留言表明您已经阅读并接受上述条款</P></td></tr></TABLE>  
</center> 
</div>
</body>