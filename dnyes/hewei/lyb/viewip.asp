<%@ language=VBScript %>
<!--#include file="setup.asp"-->
<%
if Request("menu") = "viewip" then
call viewip()
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
          <td width="75%" class="shadow">当前位置:<a href=<%=url%>><%=name%></a> - <a href=index.asp>留言首页</a> - 登陆后台</td>                    
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
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%" height="3">                    
  <tr>                    
    <td width="100%"></td>                    
  </tr>                     
  </table>                     
  </center>                 
    <table cellspacing=0 cellpadding=0 width=520 border=0 align="center">                
      <tr>                
          <td width="25"><img height=35 src="images/b.gif" width=25 border=0></td>                 
          <td width=464 background=images/b1.gif><img height=1                  
            src="images/space.gif" width=1 border=0></td>                 
          <td width="25"><img height=35 src="images/b2.gif" width=25                  
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
        <form name="reply" method="post" action="viewip.asp">   
	<input type=hidden name=id value=<%=Request("id")%>>   
	<input type=hidden name=menu value=viewip>   
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
          <td width="100%" class="shadow" align="center">请点击确定</td>   
        </tr>   
      </table>   
      </td>   
      <td width="75%" bgcolor="#F8FCF8">   
      <input type="submit" value="确定" name="ok" style="width: 50; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''"></td>   
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
   
sub viewip()   
   
password = Request.form("password")   
id = Request.form("id")   
save = Request.form("save")   
   
if id = "" then   
message=message&"对不起,你的ID号不能为空!\n"   
end if   
   
if password <> passwords then   
message=message&"您的管理口令不正确!\n"   
end if   
   
	if message <> "" then   
	call error(""&message&"")   
	Response.end   
	end if   
   
	sql = "select * from gbook where id=" & id   
	Set Rs = Server.CreateObject("ADODB.Recordset")   
	Rs.open sql,conn,1,1   
	username=Rs("username")   
	ip=Rs("ip")   
	Rs.close   
   
	if save <> "" then   
	Response.cookies("password") = password   
	Response.cookies("save") = save   
	else   
	Response.cookies("password") = ""   
	Response.cookies("save") = ""   
	end if   
   
set Rs = nothing   
conn.close   
set conn = nothing   
   
%>   
   
<script>alert('<%=username%>的IP地址:<%=ip%>');history.back();</script><script>window.close();</script>   
   
<%   
Response.End   
   
end sub   
%>