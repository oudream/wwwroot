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
          <td width="75%" class="shadow">��ǰλ��:<a href=<%=url%>><%=name%></a> - <a href=index.asp>������ҳ</a> -     
            �ظ�����</td>                    
          <td width="25%" class="shadow" align="right">����ʱ��:<%=time%></td>                    
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
{alert('�ظ����ڷ���,�����ĵȴ�!');return false;}document.reply.submit.disabled = true; this.document.reply.submit();}}  
function validate(theform) {  
	if (theform.password.value=="" || theform.reply.value=="") {  
		alert("�������ͻظ������Ǳ�����д��!");  
		clckcnt=0;  
		return false; }  
	if (postmaxchars != 0) {  
		if (theform.reply.value.length > <%=bodymax%>) {  
			alert("�����Ϣ̫����.\n\n�������� "+<%=bodymax%>+" �ֽ�����.\n��ǰ���� "+theform.reply.value.length+" �ֽ�.");  
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
          <td width="100%" class="shadow" align="center">����Ա����</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="50%">  
      <input type="password" size="20" name="password" style="width: 200; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''" value="<%=Request.cookies("password")%>"></td>  
          <td width="50%" align="center" class="shadow"><input type=hidden name=save size="20" value="<%=Request.cookies("save")%>">[<font onclick="save.value='save'" style="cursor:hand">��ס����</font>]</td>  
        </tr>  
      </table>  
      </td>  
    </tr>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">����д�ظ�</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="50%">  
      <textarea rows="3" cols="20" name="reply" style="width: 200" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''"></textarea></td>  
          <td width="50%" align="center" class="shadow">��</td>  
        </tr>  
      </table>  
      </td>  
    </tr>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">����ȷ��</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <input type="submit" value="����" name="submit" style="width: 50; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''"></td>  
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
message=message&"�Բ�����Ļظ�ID�Ų���Ϊ��!\n"  
end if  
  
sql = "select * from gbook where id="&replyid  
Set Rs = Server.CreateObject("ADODB.Recordset")  
Rs.open sql,conn,1,1  
  
if Rs.eof and Rs.bof then  
message=message&"�Բ�����Ļظ�ID�Ų�����!\n"  
end if  
Rs.close  
  
if password <> passwords then  
message=message&"���Ĺ�������ȷ!\n"  
end if  
  
if reply = "" then  
message=message&"�ظ����ݲ���Ϊ��!\n"  
end if  
  
if reply <>"" and len(reply)> bodymax then  
message=message&"�Բ��𣬻ظ��������ܳ���"&bodymax&"��,лл!\n"  
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