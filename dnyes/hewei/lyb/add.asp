<%@ language=VBScript %>
<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<!--#include file="code.asp"-->
<%
if Request("menu") = "addto" then
call addto()
else call index()
end if
sub index()
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=gbook_name%></title>
<link rel="stylesheet" href="style.css">
</head>

<body bgcolor="#CCCCCC">
  <div align=center> 
    <table cellspacing=0 cellpadding=0 width=520 border=0 align="center">
      <tbody> 
      <tr>
          <td width="25"><img height=35 src="images/c.gif" width=25 border=0></td> 
          <td width=464 background=images/c1.gif><img height=1  
            src="images/space.gif" width=1 border=0></td> 
          <td width="25"><img height=35 src="images/c2.gif" width=25  
        border=0></td> 
      </tr>
      <tr>
          <td background=images/l.gif width="25"><img height=1  
            src="images/space.gif" width=1 border=0></td> 
        <td width=464 bgcolor="#F8FCF8">  
          <div align="center">
            <center>
            <table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr>
      <td width="100%" bgcolor="#FFFFFF">
      <table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%" bgcolor="#F8FCF8">
        <tr>
          <td width="75%" class="shadow">当前位置:<a href=<%=url%>><%=name%></a> - <a href=index.asp>留言首页</a> - 添加留言</td>               
          <td width="25%" class="shadow" align="right">进入时间:<%=time%></td>               
        </tr>               
      </table>               
      </td>               
              </tr>               
            </table>               
            </center>               
          </div>               
        </td>                
          <td background=images/r.gif width="25"><img height=1                 
            src="images/space.gif" width=1 border=0></td>                
      </tr>               
      <tr>               
          <td width="25"><img height=11 src="images/bl.gif" width=25 border=0></td>                
          <td width=464 background=images/bm.gif><img height=1                 
            src="images/space.gif" width=1 border=0></td>                
          <td width="25"><img height=11 src="images/br.gif" width=25                 
        border=0></td>                
      </tr>               
      </tbody>                        
    </table>                       
<center>                
</div>               
<script language="JavaScript1.2">      
<!--      
clckcnt = 0;      
var postmaxchars = <%=bodymax%>;      
function clckcntr(eventobject){if(event.ctrlKey && window.event.keyCode==13){clckcnt++;if (clckcnt>1)       
{alert('留言正在发出,请耐心等待!');return false;}document.form.submit.disabled = true; this.document.form.submit();}}      
function validate(theform) {      
	if (theform.username.value=="" || theform.body.value=="") {      
		alert("姓名和留言是必须填写的!");      
		clckcnt=0;      
		return false; }      
	if (postmaxchars != 0) {      
		if (theform.body.value.length > <%=bodymax%>) {      
			alert("你的信息太长了.\n\n请限制在 "+<%=bodymax%>+" 字节以内.\n当前已有 "+theform.body.value.length+" 字节.");      
			clckcnt=0;      
			return false; }      
		else { document.form.submit.disabled = true; return true; }      
	} else {       
document.form.submit.disabled = true; return true; }      
}      
function checklength(theform) {      
	if (postmaxchars != 0) { message = "\n最大的字符为 "+<%=bodymax%>+" 字节."; }      
	else { message = ""; }      
	alert("你的信息已经有 "+theform.body.value.length+" 字节."+message);      
}      
-->      
</script>               
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%" height="3">               
  <tr>               
    <td width="100%"></td>               
  </tr>                
  </table>                
    <table cellspacing=0 cellpadding=0 width=520 border=0 align="center">                
      <tr>                 
          <td width="25"><img height=35 src="images/d.gif" width=25 border=0></td>                
          <td width=464 background=images/d1.gif><img height=1                 
            src="images/space.gif" width=1 border=0></td>                
          <td width="25"><img height=35 src="images/d2.gif" width=25                 
        border=0></td>                
      </tr>                
      <tr>                 
          <td background=images/l.gif width="25"><img height=1                 
            src="images/space.gif" width=1 border=0></td>                
        <td width=464 bgcolor="#F8FCF8">                 
          <div align="center">             
            <center>             
            <table border="0" width="100%" cellspacing="0" cellpadding="0">             
    <form name="form" method="post" action="add.asp" onsubmit="return validate(this)">      
    <input type="hidden" name="menu" value="addto"><input type="hidden" name="ip" value="<%=Request.ServerVariables("REMOTE_ADDR")%>">      
    <tr>      
      <td width="25%" bgcolor="#F8FCF8">      
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">      
        <tr>      
          <td width="100%" class="shadow" align="center">你的昵称(必须)</td>      
        </tr>      
      </table>      
      </td>      
      <td width="75%" bgcolor="#F8FCF8">      
      <input type="text" size="20" name="username" style="width: 200; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''"></td>      
    </tr>      
    <tr>      
      <td width="25%" bgcolor="#F8FCF8">      
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">      
        <tr>      
          <td width="100%" class="shadow" align="center">你的电邮(可选)</td>      
        </tr>      
      </table>      
      </td>      
      <td width="75%" bgcolor="#F8FCF8">      
      <input type="text" size="20" name="email" style="width: 200; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''"></td>      
    </tr>      
    <tr>      
      <td width="25%" bgcolor="#F8FCF8">      
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">      
        <tr>      
          <td width="100%" class="shadow" align="center">你的主页(可选)</td>      
        </tr>      
      </table>      
      </td>      
      <td width="75%" bgcolor="#F8FCF8">      
      <input type="text" size="20" name="homeurl" style="width: 200; height: 18" value="http://" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''"></td>      
    </tr>      
    <tr>      
      <td width="25%" bgcolor="#F8FCF8">      
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">      
        <tr>      
          <td width="100%" class="shadow" align="center">你的OICQ(可选)</td>      
        </tr>      
      </table>      
      </td>      
      <td width="75%" bgcolor="#F8FCF8">      
      <input type="text" size="20" name="qq" style="width: 200; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''"></td>      
    </tr>      
    <tr>      
      <td width="25%" bgcolor="#F8FCF8">      
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">      
        <tr>      
          <td width="100%" class="shadow" align="center">你的表情(可选)</td>      
        </tr>      
      </table>      
      </td>      
      <td width="75%" bgcolor="#F8FCF8">      
      <input type="hidden" name="brow">      
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%" height="36">      
        <tr>      
          <td width="7%" height="18" align="center"><img border="0" src="images/em01.gif" onclick="brow.value='images/em01.gif'" style="cursor:hand"></td>      
          <td width="7%" height="18" align="center"><img border="0" src="images/em02.gif" onclick="brow.value='images/em02.gif'" style="cursor:hand"></td>      
          <td width="7%" height="18" align="center"><img border="0" src="images/em03.gif" onclick="brow.value='images/em03.gif'" style="cursor:hand"></td>      
          <td width="7%" height="18" align="center"><img border="0" src="images/em04.gif" onclick="brow.value='images/em04.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em05.gif" onclick="brow.value='images/em05.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em06.gif" onclick="brow.value='images/em06.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em07.gif" onclick="brow.value='images/em07.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em08.gif" onclick="brow.value='images/em08.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em09.gif" onclick="brow.value='images/em09.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em10.gif" onclick="brow.value='images/em10.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em11.gif" onclick="brow.value='images/em11.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em12.gif" onclick="brow.value='images/em12.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em13.gif" onclick="brow.value='images/em13.gif'" style="cursor:hand"></td>      
        </tr>      
        <tr>      
          <td width="7%" height="18" align="center"><img border="0" src="images/em14.gif" onclick="brow.value='images/em14.gif'" style="cursor:hand"></td>      
          <td width="7%" height="18" align="center"><img border="0" src="images/em15.gif" onclick="brow.value='images/em15.gif'" style="cursor:hand"></td>      
          <td width="7%" height="18" align="center"><img border="0" src="images/em16.gif" onclick="brow.value='images/em16.gif'" style="cursor:hand"></td>      
          <td width="7%" height="18" align="center"><img border="0" src="images/em17.gif" onclick="brow.value='images/em17.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em18.gif" onclick="brow.value='images/em18.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em19.gif" onclick="brow.value='images/em19.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em20.gif" onclick="brow.value='images/em20.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em21.gif" onclick="brow.value='images/em21.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em22.gif" onclick="brow.value='images/em22.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em23.gif" onclick="brow.value='images/em23.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em24.gif" onclick="brow.value='images/em24.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em25.gif" onclick="brow.value='images/em25.gif'" style="cursor:hand"></td>      
          <td width="8%" height="18" align="center"><img border="0" src="images/em26.gif" onclick="brow.value='images/em26.gif'" style="cursor:hand"></td>      
        </tr>      
      </table>      
      </td>      
    </tr>      
    <tr>      
      <td width="25%" bgcolor="#F8FCF8">      
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">      
        <tr>      
          <td width="100%" class="shadow" align="center">你的头像(可选)</td>      
        </tr>      
      </table>      
      </td>      
      <td width="75%" bgcolor="#F8FCF8">      
      <input type="hidden" name="face">      
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">      
        <tr>      
          <td width="16%" height="36" align="center"><img border="0" src="images/face01.gif" onclick="face.value='images/face01.gif'" style="cursor:hand"></td>      
          <td width="16%" height="36" align="center"><img border="0" src="images/face02.gif" onclick="face.value='images/face02.gif'" style="cursor:hand"></td>      
          <td width="17%" height="36" align="center"><img border="0" src="images/face03.gif" onclick="face.value='images/face03.gif'" style="cursor:hand"></td>      
          <td width="17%" height="36" align="center"><img border="0" src="images/face04.gif" onclick="face.value='images/face04.gif'" style="cursor:hand"></td>      
          <td width="17%" height="36" align="center"><img border="0" src="images/face05.gif" onclick="face.value='images/face05.gif'" style="cursor:hand"></td>      
          <td width="17%" height="36" align="center"><img border="0" src="images/face06.gif" onclick="face.value='images/face06.gif'" style="cursor:hand"></td>      
        </tr>      
      </table>      
      </td>      
    </tr>      
    <tr>      
      <td width="25%" bgcolor="#F8FCF8">      
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">      
        <tr>      
          <td width="100%" class="shadow" align="center">你的留言(必须)</td>      
        </tr>      
      </table>      
      </td>      
      <td width="75%" bgcolor="#F8FCF8">      
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">      
        <tr>      
          <td width="50%">      
      <textarea rows="3" cols="20" name="body" style="width: 200" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''"></textarea></td>      
          <td width="50%" align="center" class="shadow">[<a href="javascript:checklength(document.form);">查看帖子长度</a>]<br>      
          [<a target="_blank" href="icode.asp">查看代码帮助</a>]</td>      
        </tr>      
      </table>      
      </td>      
    </tr>      
    <tr>      
      <td width="25%" bgcolor="#F8FCF8">      
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">      
        <tr>      
          <td width="100%" class="shadow" align="center">点击发送(必须)</td>      
        </tr>      
      </table>      
      </td>      
      <td width="75%" bgcolor="#F8FCF8">      
      <input type="submit" value="发送" name="submit" style="width: 50; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''" onclick="return clckcntr();"></td>      
    </tr>      
    </form>             
            </table>             
            </center>             
          </div>             
        </td>               
          <td background=images/r.gif width="25"><img height=1                
            src="images/space.gif" width=1 border=0></td>               
      </tr>               
      <tr>                
          <td width="25"><img height=11 src="images/bl.gif" width=25 border=0></td>               
          <td width=464 background=images/bm.gif><img height=1                
            src="images/space.gif" width=1 border=0></td>               
          <td width="25"><img height=11 src="images/br.gif" width=25                
        border=0></td>               
      </tr>               
    </table>               
</body>               
               
</html>               
<%      
      
      
Response.End      
end sub      
      
sub addto()      
dim username, email, homeurl, qq, body, addtime, sql, Rs, ip, face, brow      
      
username=htmlencode2(trim(Request.form("username")))      
email=htmlencode2(trim(Request.form("email")))      
homeurl=htmlencode2(trim(Request.form("homeurl")))      
qq=htmlencode2(trim(Request.form("qq")))      
body=icode2html(htmlencode2(Request.form("body")), false, true)      
face=Request.form("face")      
ip=Request.ServerVariables("REMOTE_ADDR")      
brow=Request.form("brow")      
      
if username = "" then      
message="您得填写您的名字哦!\n"      
end if      
      
sql="select * from retain"      
set nrs = Server.CreateObject("ADODB.Recordset")      
nrs.open sql,conn,1,1      
do while not nrs.eof      
if username <>"" and username = nrs("retain_name") then      
message=message&"您填写的名字已经被站长保留了!\n"      
end if      
nrs.movenext      
loop      
nrs.close      
set nrs = nothing      
      
if email <> "" and IsValidEmail(email)=false then      
message=message&"您的电子邮件是不是错了?\n"      
end if      
      
if qq <> "" and isInteger(qq) = false then      
message=message&"对不起，您所填写的QQ号码不是数字哦,这样是不行的!\n"      
end if      
      
if qq <> "" and len(qq) < 4 then      
message=message&"好像没有小于4位号码的QQ吧!\n"      
end if      
      
if qq <> "" and len(qq) > 11 then      
message=message&"好像还没有超过11位号码的QQ吧!\n"      
end if      
      
if homeurl = "http://" or homeurl="" then      
homeurl="http://www.anli.net"      
end if      
      
if body = "" then      
message=message&"留言内容不能为空!\n"      
end if      
      
if face = "" then      
face="images/face01.gif"      
end if      
      
if body <> "" and Len(body)> bodymax then      
message=message&"对不起,留言字数不能超过 "&bodymax&" 字,谢谢!\n"      
end if      
      
if brow = "" then      
brow="images/em01.gif"      
end if      
	if message<> "" then      
	call error(""&message&"")      
	else      
      
sql = "select * from gbook"      
Set Rs = Server.CreateObject("ADODB.Recordset")      
Rs.open sql,conn,3,2      
Rs.addnew      
Rs("username")=username      
Rs("email")=email      
Rs("homeurl")=homeurl      
Rs("qq")=qq      
Rs("body")=body      
Rs("face")=face      
Rs("brow")=brow      
Rs("ip")=ip      
Rs("addtime")=now      
Rs.update      
Rs.close      
      
	sql="select * from admin"      
	Rs.open sql,conn,3,2      
	if date <> today_time then      
	Rs("today_count") = 1      
	else      
	Rs("today_count") = Rs("today_count")+1      
	end if      
	Rs("today_time") = date      
	Rs.update      
	Rs.close      
      
set Rs = nothing      
conn.close      
set conn = nothing      
      
Response.redirect "index.asp"      
Response.End      
	end if      
end sub      
%>