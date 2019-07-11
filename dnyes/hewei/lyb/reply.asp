<%@ language=VBScript %>
<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
call index()
sub index()

if Request("menu") = "replyadd" then
call replyadd()
end if

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
              <tr>
      <td width="100%" bgcolor="#FFFFFF">
      <table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%" bgcolor="#F8FCF8">
        <tr>
          <td width="75%" class="shadow">当前位置:<a href=<%=url%>><%=name%></a> - <a href=index.asp>留言首页</a> -     
            回复公告</td>                    
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
{alert('回复正在发出,请耐心等待!');return false;}document.reply.submit.disabled = true; this.document.reply.submit();}}  
function validate(theform) {  
	if (theform.password.value=="" || theform.reply.value=="") {  
		alert("管理口令和回复内容是必须填写的!");  
		clckcnt=0;  
		return false; }  
	if (postmaxchars != 0) {  
		if (theform.reply.value.length > <%=bodymax%>) {  
			alert("你的信息太长了.\n\n请限制在 "+<%=bodymax%>+" 字节以内.\n当前已有 "+theform.reply.value.length+" 字节.");  
			clckcnt=0;  
			return false; }  
		else { document.reply.submit.disabled = true; return true; }  
	} else {   
document.reply.submit.disabled = true; return true; }  
}  
-->  
</script>                   
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%" height="3">                    
  <tr>                    
    <td width="100%"></td>                    
  </tr>                     
  </table>                     
  </center>                 
    <table cellspacing=0 cellpadding=0 width=520 border=0 align="center">                
      <tr>                
          <td width="25"><img height=35 src="images/g.gif" width=25 border=0></td>                 
          <td width=464 background=images/g1.gif><img height=1                  
            src="images/space.gif" width=1 border=0></td>                 
          <td width="25"><img height=35 src="images/g2.gif" width=25                  
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
        <form name="reply" method="post" action="reply.asp" onsubmit="return validate(this)">  
	<input type=hidden name=menu value=replyadd>  
	<input type=hidden name=replyid value=<%=request("replyid")%>>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">管理员密码</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="50%">  
      <input type="password" size="20" name="password" style="width: 200; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''" value="<%=Request.cookies("password")%>"></td>  
          <td width="50%" align="center" class="shadow"><input type=hidden name=save size="20" value="<%=Request.cookies("save")%>">[<font onclick="save.value='save'" style="cursor:hand">记住密码</font>]</td>  
        </tr>  
      </table>  
      </td>  
    </tr>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">这里写回复</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="50%">  
      <textarea rows="3" cols="20" name="reply" style="width: 200" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''"></textarea></td>  
          <td width="50%" align="center" class="shadow">　</td>  
        </tr>  
      </table>  
      </td>  
    </tr>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">请点击确定</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <input type="submit" value="发送" name="submit" style="width: 50; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''"></td>  
    </tr>  
    </form>               
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
    </table>                           
</body>                   
                   
</html>                   
<%  
  
  
Response.End  
end sub  
  
sub replyadd()  
dim password,reply,replyid  
  
password = Request.form("password")  
reply = htmlencode2(trim(Request.form("reply")))  
replyid = Request.form("replyid")  
save = Request.form("save")  
  
if replyid = "" then  
message=message&"对不起，你的回复ID号不能为空!\n"  
end if  
  
sql = "select * from gbook where id="&replyid  
Set Rs = Server.CreateObject("ADODB.Recordset")  
Rs.open sql,conn,1,1  
  
if Rs.eof and Rs.bof then  
message=message&"对不起，你的回复ID号不存在!\n"  
end if  
Rs.close  
  
if password <> passwords then  
message=message&"您的管理口令不正确!\n"  
end if  
  
if reply = "" then  
message=message&"回复内容不能为空!\n"  
end if  
  
if reply <>"" and len(reply)> bodymax then  
message=message&"对不起，回复字数不能超过"&bodymax&"字,谢谢!\n"  
end if  
  
	if message <> "" then  
	call error(""&message&"")  
	Response.end  
  
	else  
  
	if save <> "" then  
	Response.cookies("password") = password  
	Response.cookies("save") = save  
	else  
	Response.cookies("password") = ""  
	Response.cookies("save") = ""  
	end if  
  
sql = "select * from reply"  
Rs.open sql,conn,3,2  
Rs.addnew  
Rs("reply")=reply  
Rs("replytime")=now  
Rs("replyid")=replyid  
Rs.update  
Rs.close  
conn.close  
set Rs = nothing  
set conn = nothing  
Response.redirect "index.asp"  
Response.End  
  
	end if  
end sub  
%>