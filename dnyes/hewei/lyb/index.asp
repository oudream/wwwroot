<%@ language=VBScript %>
<!--#include file="setup.asp"-->
<!--#include file="online.asp"-->
<%
	on error resume next

	Set Rs1 = Server.CreateObject("ADODB.Recordset")
	sql1="select * from admin"
	Rs1.open sql,conn,3,2
	if date <> today_time then
	Rs1("today_count") = 0
	end if
	Rs1("today_time") = date
	Rs1.update
	Rs1.close
	set Rs1 = nothing

sql="select * from gbook order by id desc"
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.open sql,conn,1,1
Rs.pagesize=pagesize
page=Request("page")
if (page-Rs.pagecount) > 0 then
page=Rs.pagecount
elseif page = "" or page < 1 then
page = 1
end if
Rs.absolutepage=page
%>
<html>

<head>
<%
sql = "select * from news order by id desc"
set srs = Server.CreateObject("ADODB.Recordset")
srs.open sql,conn,1,1
do while not srs.eof
news=news&""""&srs("news")&" ["&srs("date")&"]"","""","
srs.movenext
loop
srs.close
set srs = nothing
%>
<script language="JavaScript1.2" type="text/javascript">
NS4 = (document.layers);
IE4 = (document.all);
   FDRboxWid = 450;
   FDRboxHgt = 18;
   FDRborWid = 2;
   FDRborCol = "#EFEFEF";
   FDRborSty = "solid";
   FDRbackCol = "#EFEFEF";
   FDRboxPad = 2;

   FDRtxtAln = "center";
   FDRlinHgt = "9pt";
   FDRfntFam = "宋体";
   FDRfntCol = "#4F4F4F";
   FDRfntSiz = "9pt";
   FDRfntWgh = "";
   FDRfntSty = "normal";
   FDRlnkCol = "#4F4F4F";
   FDRlnkDec = "none";
   FDRhovCol = "#000000";

   FDRgifSrc = "";
   FDRgifInt = 60;

   FDRblendInt = 5;
   FDRblendDur = 1;
   FDRmaxLoops = 100;

   FDRendWithFirst = true;
   FDRreplayOnClick = true;
   
   FDRjustFlip = false;
   FDRhdlineCount = 0;
</script>
<script language='JavaScript1.2' type='text/javascript'>
prefix="";
arNews = [
<%=news%>
"<%=gbook_name%> 公告栏",""
]
</script>
<script language='JavaScript1.2' src='fader.js' type='text/javascript'></script>
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
          <td width="75%" class="shadow">当前位置:<a href=<%=url%>><%=name%></a>  
            - <a href=index.asp>留言首页</a>              
            - 浏览留言</td>             
          <td width="25%" class="shadow" align="right">进入时间:<%=time%></td>             
        </tr>             
      </table>             
      </td>             
              </tr>             
              <tr>             
                <td width="100%">             
                  <div align="center">             
                    <center>             
                    <table border="0" width="100%" cellspacing="0" cellpadding="0">             
                      <tr>             
                            <td width="15%" align="center"><a href="login.asp"><img border="0" src="images/admin.gif"></a></td>
                            <td width="15%" align="center">&nbsp;</td>             
          <td width="40%">             
          <table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber4">             
            <form action="search.asp" method="post">   
            <tr>   
              <td width="100%" class="shadow" align="center">共<%=Rs.recordcount%>条留言 其中<%=today_count%>条是今天的</td>   
            </tr>   
            <tr>   
              <td width="80%" align="center">   
              <input type="text" size="20" name="keyword" style="width: 100; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''">   
              <input type="submit" value="搜索" name="ok" style="width: 50; height: 18" onmouseover="this.style.backgroundColor='#FFFFFF'" onmouseout="this.style.backgroundColor=''"></td>   
            </tr>   
            </form>              
            </table>              
          </td>              
                            <td width="15%" align="center"> <a href="http://www.sczz.com/book/register.asp" target="_blank"> 
                              </a><a href="index.asp"><img border="0" src="images/read.gif"></a></td>
                            <td width="15%" align="center"><a href="add.asp"><img border="0" src="images/new.gif"></a></td>             
                      </tr>             
                    </table>             
                    </center>             
                  </div>             
                </td>             
              </tr>             
              <tr>             
    <td width="100%" bgcolor="#F8FCF8" align="center" valign="top" height="3">　</td>             
              </tr>             
              <tr>             
    <td width="100%" bgcolor="#F8FCF8" align="center" valign="top" height="20"><div id="elFader" style="position:relative;visibility:hidden;width:100%"></div></td>             
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
</div>             
<%  
do while not Rs.eof  
i = i + 1  
if i > Rs.pagesize then  
exit do  
end if  
ips=split(rs("ip"),".")  
ip=""&ips(0)&"."&ips(1)&".*.*"  
%>             
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%" height="3">             
  <tr>             
    <td width="100%"></td>            
    </tr>            
  </table>            
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
            <table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolor="#F8FCF8">           
              <tr>           
      <td width="132" bgcolor="#F8FCF8" valign="top">           
      <table border="0" cellpadding="2" cellspacing="0" style="table-layout: fixed" bordercolor="#111111" width="100%" style="">           
        <tr>           
          <td width="100%" align="center" class="shadow" height="5">           
          </td>           
        </tr>           
        <tr>           
          <td width="100%" align="center" class="shadow">           
          <img border="0" src="<%=Rs("face")%>"></td>           
        </tr>           
        <tr>           
          <td width="100%" class="shadow" height="5"></td>          
        </tr>          
        <tr>           
          <td width="100%" class="shadow">&nbsp; 大名:<%=Rs("username")%></td>           
        </tr>           
        <tr>           
          <td width="100%" class="shadow">&nbsp; OICQ:<%if Rs("qq") <> "" then%><%=Rs("qq")%><%else%>没有OICQ<%end if%></td>           
        </tr>           
      </table>           
      </td>           
      <td width="362" bgcolor="#F8FCF8">           
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber6">           
        <tr>           
          <td width="100%" bgcolor="#EFEFEF">           
          <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber7">           
            <tr>           
              <td width="100%" class="shadow">第 <%=Rs.recordcount-Rs.pagesize*(page-1)-i+1%> 条留言 是 <%=Rs("username")%> 在 <%=Rs("addtime")%> 发表的</td>             
            </tr>             
          </table>             
          </td>             
        </tr>             
        <tr>             
          <td width="100%" style="border-left-width: 1; border-right-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1">             
          <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber8" height="40" style="table-layout: fixed">             
            <tr>             
              <td width="8%" class="shadow" valign="top" align="center">             
              <img border="0" src="<%=Rs("brow")%>"></td>             
              <td width="92%" class="shadow" valign="top"><%=Rs("body")%>             
<%   
sql = "select * from reply where replyid="& Rs("id") &" order by id desc"   
set Rss = Server.CreateObject("ADODB.Recordset")   
Rss.open sql,conn,1,1   
do while not Rss.eof   
%>             
<br><hr noshade size="1" style="border-style: dotted"><font color=#999999>站长回复:<%=Rss("reply")%><br>回复时间:(<%=Rss("replytime")%>)</font><br>             
<%   
Rss.movenext   
loop   
Rss.close   
set Rss = nothing   
%></td>             
            </tr>             
          </table>             
          </td>             
        </tr>             
        <tr>             
          <td width="100%" bgcolor="#EFEFEF">             
          <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">             
            <tr>             
              <td width="18%" align="center" class="shadow"><%if Rs("email") <> "" then%><a href='mailto:<%=Rs("email")%>'>电子邮件</a><%else%>没有信箱<%end if%></td>   
              <td width="18%" align="center" class="shadow"><a href='<%=Rs("homeurl")%>' target=_blank>个人主页</a></td>   
              <td width="18%" align="center" class="shadow"><%if Rs("qq") <> "" then%><a href='http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<%=Rs("qq")%>'>腾讯OICQ<%else%>没有OICQ<%end if%></a></td>   
              <td width="16%" align="center">　</td> 
              <td width="10%" align="center" class="shadow"><a href=viewip.asp?id=<%=Rs("id")%>>查IP</a></td> 
              <td width="10%" align="center" class="shadow"><a href=reply.asp?replyid=<%=Rs("id")%>>回复</a></td> 
              <td width="10%" align="center" class="shadow"><a href=del.asp?id=<%=Rs("id")%>>删除</a></td>  
            </tr>  
          </table>  
          </td>  
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
    </table>  
<% 
Rs.movenext 
loop 
for i=1 to count 
if i<>1 and (i-1)/6=int((i-1)/6) then 
guest=guest&"</tr><tr>" 
end if 
guest=guest&"<td width=100><img border=0 src=images/guest.gif> 留言客"&i&"号</td>" 
next 
if count/6<>int(i/6) then 
for j=1 to (int(count/6)+1)*6-count 
guest=guest&"<td width=100>　</td>" 
next 
end if 
%>  
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#4F4F4F" width="100%" height="3">  
  <tr>  
    <td width="100%"></td>  
  </tr>  
</table>  
<div align="center">  
  <center>  
    <table cellspacing=0 cellpadding=0 width=520 border=0 align="center">  
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
            <table border="0" width="100%" cellspacing="0" cellpadding="0" height="66">  
              <tr>  
      <td width="100%" bgcolor="#EFEFEF" class="shadow" height="19">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber11" height="17">  
        <tr>  
          <td width="100%" class="shadow" height="17">目前留言本共有<%=count%>人在线</td>  
        </tr>  
      </table>  
      </td>  
              </tr>  
              <tr>  
      <td width="100%" bgcolor="#111111" align="center" height="1">  
      <img height=1       
            src="images/space.gif" width=1 border=0></td>  
              </tr>  
              <tr>  
      <td width="100%" bgcolor="#F8FCF8" align="center" height="14">  
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber11">  
        <tr>  
          <%=guest%>  
        </tr>  
      </table></td>  
              </tr>  
              <tr>  
                <td width="100%" height="14"></td>  
              </tr>  
              <tr>  
      <td width="100%" class="white" height="14">共有<%=Rs.pagecount%>页 你正在浏览第<%=page%>页 跳到第<%for go=1 to Rs.pagecount%> <a href=index.asp?page=<%=go%>><%=go%></a> <%next%>页<% if page<>1 then %> <a href=index.asp?page=<%=page-1%>>上一页</a><%end if%><% if Rs.pagecount-page <> 0 then %> <a href=index.asp?page=<%=page+1%>>下一页</a><%end if%></td>   
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
  <table border="0" width="100%" cellspacing="0" cellpadding="0">   
    <tr>   
      <td width="100%"></td>   
    </tr>   
  </table>   
</body>   
   
</html>   
