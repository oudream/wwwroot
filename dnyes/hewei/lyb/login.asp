<%@ language=VBScript %>
<!--#include file="setup.asp"-->
<%
if Request("menu") = "login" then
	call login()
elseif Request("menu") = "logout" then
	call logout()
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
          <td width="25"><img height=35 src="images/i.gif" width=25 border=0></td> 
          <td width=464 background=images/i1.gif><img height=1  
            src="images/space.gif" width=1 border=0></td> 
          <td width="25"><img height=35 src="images/i2.gif" width=25  
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
          <td width="75%" class="shadow">��ǰλ��:<a href=<%=url%>><%=name%></a> - <a href=index.asp>������ҳ</a> - ��½��̨</td>                   
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
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%" height="3">                   
  <tr>                   
    <td width="100%"></td>                   
  </tr>                    
  </table>                    
  </center>                
    <table cellspacing=0 cellpadding=0 width=520 border=0 align="center">               
      <tr>               
          <td width="25"><img height=35 src="images/h.gif" width=25 border=0></td>                
          <td width=464 background=images/h1.gif><img height=1                 
            src="images/space.gif" width=1 border=0></td>                
          <td width="25"><img height=35 src="images/h2.gif" width=25                 
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
        <form name="reply" method="post" action="login.asp">  
    <input type=hidden name=menu value=login>  
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
          <td width="100%" class="shadow" align="center">����ȷ��</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <input type="submit" value="ȷ��" name="ok" style="width: 50; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''"></td>  
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
  
  
end sub  
  
sub login()  
  
	password = Request.form("password")  
	save = Request.form("save")  
	if password <> passwords then  
	message=message&"���Ĺ�������ȷ!\n"  
	end if  
  
	if message <> "" then  
	call error(""&message&"")  
	Response.end  
	end if  
  
	if save <> "" then  
	Response.cookies("password") = password  
	Response.cookies("save") = save  
	else  
	Response.cookies("password") = ""  
	Response.cookies("save") = ""  
	end if  
  
	Response.cookies("admin") = password  
	Response.redirect "administer.asp"  
  
end sub  
  
sub logout()  
  
	Response.cookies("admin") = ""  
	Response.cookies("password") = ""  
	Response.cookies("save") = ""  
  
	Response.redirect "index.asp"  
end sub  
%>