<%@ language=VBScript %>
<!--#include file="setup.asp"-->
<%
if Request("menu") = "news" then
call news()
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
          <td width="75%" class="shadow">��ǰλ��:<a href=<%=url%>><%=name%></a> - <a href=index.asp>������ҳ</a> -    
            ��������</td>                   
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
        <form name="form" method="post" action="news.asp"> 
    <input type=hidden name=menu value=news><input type=hidden name=date size="20" value="<%=now%>"> 
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
          <td width="100%" class="shadow" align="center">�����¹���</td> 
        </tr> 
      </table> 
      </td> 
      <td width="75%" bgcolor="#F8FCF8"> 
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%"> 
        <tr> 
          <td width="100%">       
          <input type="text" size="255" name="news" style="width: 200; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''"></td> 
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
 
 
Response.End 
end sub 
 
sub news() 
 
password = Request.form("password") 
newsinfo = Request.form("news") 
newsdate = Request.form("date") 
save = Request.form("save") 
 
if newsinfo = "" then 
message=message&"�Բ���,�������ݲ���Ϊ��!\n" 
end if 
 
if password <> passwords then 
message=message&"���Ĺ�������ȷ!\n" 
end if 
on error resume next 
if message <> "" then 
	call error(""&message&"") 
	Response.end 
	else 
 
    sql = "select * from news" 
    Set Rs = Server.CreateObject("ADODB.Recordset") 
    Rs.open sql,conn,3,2 
    Rs.addnew 
    Rs("news")=newsinfo 
    Rs("date")=newsdate 
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