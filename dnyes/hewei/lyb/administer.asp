<%@ language=VBScript %>
<!--#include file="setup.asp"-->
<%
if Request.cookies("admin") = "" then
message="对不起,非法登陆!\n"
call error(""&message&"")
Response.end
end if
if Request("menu") = "setup" then
call setup()
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
          <td width="25"><img height=35 src="images/e.gif" width=25 border=0></td> 
          <td width=464 background=images/e1.gif><img height=1  
            src="images/space.gif" width=1 border=0></td> 
          <td width="25"><img height=35 src="images/e2.gif" width=25  
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
            管理设置</td>                  
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
        <form name="form" method="post" action="administer.asp">  
    <input type=hidden name=menu value=setup>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">留言本名称</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <input type="text" size="20" name="gbook_name" style="width: 200; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''" value="<%=gbook_name%>"></td>  
    </tr>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">管理员密码</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <input type="password" size="20" name="password" style="width: 200; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''" value="<%=passwords%>"></td>  
    </tr>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">主页的名称</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <input type="text" size="20" name="name" style="width: 200; height: 18" value="<%=name%>" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''"></td>  
    </tr>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">网站的地址</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <input type="text" size="20" name="url" style="width: 200; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''" value="<%=url%>"></td>  
    </tr>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">长度的限制</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <input type="text" size="20" name="bodymax" style="width: 200; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''" value="<%=bodymax%>"></td>  
    </tr>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">每页留言数</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <input type="text" size="20" name="pagesize" style="width: 200; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''" value="<%=pagesize%>"></td>  
    </tr>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">你的IP地址</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <input type="text" size="20" style="width: 200; height: 18" value="<%=Request.ServerVariables("REMOTE_ADDR")%>" disabled="true"></td>  
    </tr>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">上次IP地址</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <input type="text" size="20" style="width: 200; height: 18" value="<%=setup_ip%>" disabled="true"></td>  
    </tr>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">最后修改在</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
      <input type="text" size="20" style="width: 200; height: 18" value="<%=setup_time%>" disabled="true"></td>  
    </tr>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">特殊用户名</td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
<%  
sql="select * from retain"  
set nrs = Server.CreateObject("ADODB.Recordset")  
nrs.open sql,conn,1,1  
do while not nrs.eof  
retain_name = retain_name & nrs("retain_name") & "|"  
nrs.movenext  
loop  
nrs.close  
set nrs = nothing  
%>  
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">  
  <tr>  
    <td width="50%">  
      <input type="text" size="20" style="width: 200; height: 18" value="<%=retain_name%>" disabled="true" name="1"></td>  
    <td width="25%" align="center">[<a href="retain.asp">添加</a>]</td>  
    <td width="25%" align="center">[<a href="delretain.asp">删除</a>]</td>  
  </tr>  
</table>  
      </td>  
    </tr>  
    <tr>  
      <td width="25%" bgcolor="#F8FCF8">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%">  
        <tr>  
          <td width="100%" class="shadow" align="center">站长公告栏<br>  
          <a href="news.asp">添加新公告</a></td>  
        </tr>  
      </table>  
      </td>  
      <td width="75%" bgcolor="#F8FCF8">  
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">  
<%  
sql="select * from news order by id desc"  
set srs = Server.CreateObject("ADODB.Recordset")  
srs.open sql,conn,1,1  
do while not srs.eof  
%>  
  <tr>  
    <td width="50%">  
      <input type="text" size="20" style="width: 200; height: 18" value="<%=srs("news")%>" disabled="true" name="1"></td>  
    <td width="25%" align="center">[<a href="modnews.asp?id=<%=srs("id")%>&oldnews=<%=srs("news")%>">修改</a>]</td>  
    <td width="25%" align="center">[<a href="delnews.asp?id=<%=srs("id")%>">删除</a>]</td>  
  </tr>  
<%  
srs.movenext  
loop  
srs.close  
set srs = nothing  
%>  
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
  
sub setup()  
  
gbook_name=Server.htmlencode(Request.form("gbook_name"))  
password=Request.form("password")  
name=Server.htmlencode(Request.form("name"))  
url=Request.form("url")  
bodymax=Request.form("bodymax")  
pagesize=Request.form("pagesize")  
setup_ip=Request.ServerVariables("REMOTE_ADDR")  
  
if gbook_name = "" then  
message="留言本名称不能为空!\n"  
end if  
  
if password = "" then  
message=message&"管理员口令不能为空!\n"  
end if  
  
if name = "" then  
message=message&"您的网站名称必须填写!\n"  
end if  
  
if url = "" then  
message=message&"您的网站地址必须填写!\n"  
end if  
  
if bodymax = "" then  
message=message&"您不能把留言内容大小限制设置为空!\n"  
end if  
  
if pagesize = "" then  
message=message&"每页显示的留言数必须填写!\n"  
end if  
  
	if message<> "" then  
	call error(""&message&"")  
	Response.end  
	end if  
  
sql = "select * from admin"  
Set Rs = Server.CreateObject("ADODB.Recordset")  
Rs.open sql,conn,3,3  
Rs("gbook_name")=gbook_name  
Rs("password")=password  
Rs("name")=name  
Rs("url")=url  
Rs("bodymax")=bodymax  
Rs("pagesize")=pagesize  
Rs("setup_ip")=setup_ip  
Rs("setup_time")=now  
Rs.update  
  
Rs.close  
set Rs = nothing  
conn.close  
set conn = nothing  
  
Response.redirect "administer.asp"  
Response.End  
  
end sub  
%>