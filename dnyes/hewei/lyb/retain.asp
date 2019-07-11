<%@ language=VBScript %>
<!--#include file="setup.asp"-->
<%
if Request("menu") = "retain" then
call retain()
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
          <td width="25"><img height=35 src="images/a.gif" width=25 border=0></td> 
          <td width=464 background=images/a1.gif><img height=1  
            src="images/space.gif" width=1 border=0></td> 
          <td width="25"><img height=35 src="images/a2.gif" width=25  
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
          <td width="75%" class="shadow">当前位置:<a href=<%=url%>><%=name%></a> - <a href=index.asp>留言首页</a> - 保留用户</td>                     
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
          <td width="25"><img height=35 src="images/j.gif" width=25 border=0></td>                  
          <td width=464 background=images/j1.gif><img height=1                   
            src="images/space.gif" width=1 border=0></td>                  
          <td width="25"><img height=35 src="images/j2.gif" width=25                   
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
        <form name="form" method="post" action="retain.asp">   
    <input type=hidden name=menu value=retain>   
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
          <td width="100%" class="shadow" align="center">保留用户名</td>   
        </tr>   
      </table>   
      </td>   
      <td width="75%" bgcolor="#F8FCF8">   
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">   
        <tr>   
          <td width="100%">         
          <input type="text" size="20" name="retain_name" style="width: 200; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''"></td>   
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
   
sub retain()   
   
password = Request.form("password")   
retain_name = Request.form("retain_name")   
save = Request.form("save")   
   
if retain_name = "" then   
message=message&"对不起,保留用户名不能为空!\n"   
end if   
   
if password <> passwords then   
message=message&"您的管理口令不正确!\n"   
end if   
on error resume next   
if message <> "" then   
	call error(""&message&"")   
	Response.end   
	else   
   
    sql = "select * from retain"   
    Set Rs = Server.CreateObject("ADODB.Recordset")   
    Rs.open sql,conn,3,2   
    Rs.addnew   
    Rs("retain_name")=retain_name   
	Rs.update   
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
   
Response.redirect "administer.asp"   
Response.End   
end if   
end sub   
%>