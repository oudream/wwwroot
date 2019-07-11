<!--#include file=config.asp -->
<%
dim getplace,getshow,adsrs,adssql,adsrsp,adssqlp,adsrss,adssqls,getip,getggwlxsz,getggwhei,getggwwid
'response.write Ggwlxsz(cint(request.querystring("i")))
getplace=cint(request.querystring("i"))

dim GaoAndKuan
adsconn.open adsdata

set adsrs1=server.createobject("adodb.recordset")
adsrs1.open "select * from place where place="&getplace,adsconn,1,1
if not adsrs1.eof then
getggwlxsz=adsrs1(2)
else
getggwlxsz=0
end if
getggwhei=adsrs1(3)
getggwwid=adsrs1(4)

GaoAndKuan=""

if getggwhei<>"" then GaoAndKuan=" height="&getggwhei&" "
if getggwwid<>"" then GaoAndKuan=GaoAndKuan&" width="&getggwwid&" "

adsrs1.close
Set adsrs1=nothing

'''''''''''''''''''''''''''''''''''''''''''''''每次显示广告位前，检测其中的各广告条是否过期，并更新状态''''''''''''''''''''''''''''''''
set adsrsp=server.createobject("adodb.recordset")
adssqlp="Select * from [ads] where act=1 and class <> 0 and  place="&getplace&" order by time"
adsrsp.open adssqlp,adsconn,1,3

while not adsrsp.eof

advertvirtualvalue=0

if adsrsp("class")=1 then
if adsrsp("click")>=adsrsp("clicks") then
advertvirtualvalue=1
end if

elseif adsrsp("class")=2 then
if adsrsp("show")>=adsrsp("shows") then
advertvirtualvalue=1
end if

elseif adsrsp("class")=3 then
if now()>=adsrsp("lasttime") then
advertvirtualvalue=1
end if

elseif adsrsp("class")=4 then
if adsrsp("click")>=adsrsp("clicks") then
advertvirtualvalue=1
end if
if adsrsp("show")>=adsrsp("shows") then
advertvirtualvalue=1
end if

elseif adsrsp("class")=5 then
if adsrsp("click")>=adsrsp("clicks") then
advertvirtualvalue=1
end if
if now()>=adsrsp("lasttime") then
advertvirtualvalue=1
end if

elseif adsrsp("class")=6 then
if adsrsp("show")>=adsrsp("shows") then
advertvirtualvalue=1
end if
if now()>=adsrsp("lasttime") then
advertvirtualvalue=1
end if

elseif adsrsp("class")=7 then
if adsrsp("click")>=adsrsp("clicks") then
advertvirtualvalue=1
end if
if adsrsp("show")>=adsrsp("shows") then
advertvirtualvalue=1
end if
if now()>=adsrsp("lasttime") then
advertvirtualvalue=1
end if
end if

if advertvirtualvalue>=1 then
adsrsp("act")=2
adsrsp.update
end if
adsrsp.movenext
wend
adsrsp.close
set adsrsp=nothing 
adsconn.close
'''''''''''''''''''''''''''''''''''''''''''''''结束 检测、更新''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''打开数据库连接
adsconn.open adsdata
set adsrs=server.createobject("adodb.recordset")
''''''''''''''''''''''''''''''''''''''''根据显示方式的不同进行显示''''''''''''''''''''''''
Select Case getggwlxsz

       Case 1 
       
       adssql="Select top 1 id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei from [ads] where act=1 and place="&getplace&" order by time"
       adsrs.open adssql,adsconn,1,3
       Call DggtXs()
       adsrs.close

       Case 2 
       
       adssql="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei from [ads] where act=1 and place="&getplace&" order by time"
       adsrs.open adssql,adsconn,1,3
       do while not adsrs.eof 
       Call DggtXs()
       adsrs.movenext
       response.write "document.write('<br><br>');"
       loop
       adsrs.close
       
       Case 3 
       
       adssql="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei from [ads] where act=1 and place="&getplace&" order by time"
       adsrs.open adssql,adsconn,1,3
       do while not adsrs.eof 
       Call DggtXs()
       adsrs.movenext
       response.write "document.write('&nbsp;&nbsp;');"
       loop
       adsrs.close

       Case 4 
       
       adssql="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei from [ads] where act=1 and place="&getplace&" order by time"
       adsrs.open adssql,adsconn,1,3
       response.write "document.write('<marquee  direction=up"&GaoAndKuan&">');"
       do while not adsrs.eof
       Call DggtXs()
       adsrs.movenext
       response.write "document.write('<br><br>'); "
       loop
       response.write "document.write('</marquee>');"
       adsrs.close 

       Case 5 
       
       adssql="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei from [ads] where act=1 and place="&getplace&" order by time"
       adsrs.open adssql,adsconn,1,3

       
       response.write "document.write('<marquee"&GaoAndKuan&">');"
       do while not adsrs.eof
       Call DggtXs()
       adsrs.movenext
       response.write "document.write('&nbsp;&nbsp;');"
       loop
       response.write "document.write('</marquee>');"
       adsrs.close 

       Case 6 
       
       adssql="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei from [ads] where act=1 and place="&getplace&" order by time"
       adsrs.open adssql,adsconn,1,3
       response.write "document.write('<marquee"&GaoAndKuan&">');"
       do while not adsrs.eof
       Call DggtXs()
       adsrs.movenext
       response.write "document.write('&nbsp;&nbsp;');"
       loop
       response.write "document.write('</marquee>');"
       adsrs.close 

       Case 7 
       dim gao1,kuan1
       
       adssql="Select top 1 id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei from [ads] where act=1 and place="&getplace&" order by time"
       adsrs.open adssql,adsconn,1,3
       
       response.write ""       
       
       gao1=""
       kuan1=""

       if adsrs("hei")<>"" then gao1=" height="&adsrs("hei")&" "
       if adsrs("wid")<>"" then kuan1=" width="&adsrs("wid")&" "
       
       response.write "window.open('ad1.asp?i="&adsrs("id")&"','忠网广告','"&kuan1&","&gao1&"');"
       response.write ""
       adsrs.close 
       
End Select 
       ''''''''''''''''''''''''''''''''''''''''''''''''断开数据连接
       
set adsrs=nothing
adsconn.close
set adsconn=nothing %>
<% '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' ''''''''''''''''''''''''''''显示单个广告条 '''''''''''''''''''''''''''''''''''''''''''''' 

Sub DggtXs() 
adsrs("show")=adsrs("show")+1
adsrs("time")=now()
adsrs.Update
if adsrs("window")=0 then
ttarg = "_blank"
else 
ttarg="_top" 
end if


if isnumeric(adsrs("hei")) then
GaoAndKuan=" height="&adsrs("hei")&" "
else

if right(adsrs("hei"),1)="%" then
if isnumeric(Left(len(adsrs("hei"))-1))=true then
 GaoAndKuan=" height="&adsrs("hei")&" "
end if
end if

end if


if isnumeric(adsrs("wid")) then
GaoAndKuan=GaoAndKuan&" width="&adsrs("wid")&" "
else
if right(adsrs("wid"),1)="%" then
if isnumeric(Left(len(adsrs("wid"))-1))=true then 
GaoAndKuan=GaoAndKuan&" width="&adsrs("wid")&" "
end if
end if
end if

 %> 
<% Select Case adsrs("xslei")
   Case "txt"%>document.write('<a title=\"<%=adsrs("sitename")%>\"  href=\"<%=DqUrl%>/url.asp?id=<%=adsrs("id")%>\" target=\"<%=ttarg%>\"><%=UBBCode(adsrs("intro"))%></a>');
<% Case "gif"%>document.write('<a href=\"<%=DqUrl%>/url.asp?id=<%=adsrs("id")%>\" target=\"<%=ttarg%>\"><img  art=\"<%=adsrs("sitename")%>\"  border=0 <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>"></a>');
<% Case "swf"%>document.write('<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http:/download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0";  <%=GaoAndKuan%>><param name=movie value="<%=adsrs("gif_url")%>"><param name=quality value=high><param name="wmode" value="transparent"></object>');
<% Case "dai"%>document.write('<iframe marginwidth=0 marginheight=0  frameborder=0 bordercolor=000000 scrolling=no  name=\"忠网WEB广告管理系统 zon.cn\" src=\"<%=DqUrl%>/daima.asp?id=<%=adsrs("id")%>\"  <%=GaoAndKuan%> ></iframe>');
<% Case else%>document.write('<a href=\"<%=DqUrl%>/url.asp?id=<%=adsrs("id")%>\" target=\"<%=ttarg%>\"><img art=\"<%=adsrs("sitename")%>\"  border=0 <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>"></a>');
<%End Select

getip=request.ServerVariables("REMOTE_ADDR")
set adsrss=server.createobject("adodb.recordset")
adssqls="select * from iplist"
adsrss.open adssqls,adsconn,1,3
adsrss.AddNew
adsrss("adid") =adsrs("id")
adsrss("time") = now()
adsrss("ip") = getip
adsrss("class") = 1
adsrss.update
adsrss.close
set adsrss=nothing


End Sub
%>