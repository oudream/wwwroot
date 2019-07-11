<!--#include file=config.asp -->
<html><head>
<meta http-equiv=Content-Language content=zh-cn>
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<title>&#24544;&#32593;.&#24191;&#21578;&#31649;&#29702;&#31995;&#32479;</title>
<meta http-equiv=Pragma content=no-cache>
<link href=style.css rel=STYLESHEET type=text/css>
<base target=_top>
<style>
<!--
		table {WORD-BREAK: break-all;}
-->
</style>
</head><body marginwidth=0 marginheight=0>
<div align=center><center><table width="100%" height="100%"><tr><td align=center>
<%
if session("masterlogin")<>"superadvertadmin" then
%>
<br><br>
<li>&#26435;&#38480;&#19981;&#36275;&#65292;&#35831;&#30331;&#38470;&#31649;&#29702;&#21518;&#36827;&#20837;&#12290;
<%
else

adsconn.open adsdata
dim getid,adsrs,adssql
getid=cint(request.querystring("id"))


Select Case request.querystring("job")
    case "close"

   set adsrs=server.createobject("adodb.recordset")
   adssql="Select id,sitename,act from [ads] where id="&getid
   adsrs.open adssql,adsconn,1,3
   adsrs("act")=0
   adsrs.Update
%>
<br><br>&#20197;&#19979;&#24191;&#21578;&#26465;&#34987;&#26242;&#20572;&#65306;<br><br>
<%=adsrs("sitename")%><br><br>

<%
adsrs.close

    case "delit"
%>
<br><br>&#21024;&#38500;&#27492;&#24191;&#21578;&#65292;&#21017;&#20854; IP &#35760;&#24405;&#20063;&#23558;&#34987;&#21024;&#38500;&#65281;<br>
&#24191;&#21578;&#21450;&#20854;IP&#35760;&#24405;&#34987;&#21024;&#38500;&#21518;&#19981;&#33021;&#24674;&#22797;&#12290;<br><br><br>
<a href=option.asp?id=<%=getid%>&job=del>&#32487;&#32493;</a>&#36824;&#26159;
<a href=javascript:window.close()>&#36864;&#20986;</a>&#65311;

<%

    case "del"
    
    adssql="delete from [ads] where id="&getid
    adsconn.execute(adssql)
    dim adssqldelip
    adssqldelip="delete from [iplist] where adid="&getid
    adsconn.execute(adssqldelip)
%>
<br><br>&#24191;&#21578;&#26465;&#34987;&#25104;&#21151;&#21024;&#38500;&#65281;<br><br>


<%
    case "yulan"
set adsrs=server.createobject("adodb.recordset")
adssql="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei from [ads] where id="&getid
adsrs.open adssql,adsconn,3,3

adsrs("show")=adsrs("show")+1
adsrs("time")=now()
adsrs.Update


if adsrs("window")=0 then
ttarg = "_blank"
else
ttarg="_top"
end if

Dim GaoAndKuan
GaoAndKuan=""

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





           Select Case adsrs("xslei")
    
            Case "txt"%><a  title="<%=adsrs("sitename")%>"  href="<%=DqUrl%>/url.asp?id=<%=adsrs("id")%>" target="<%=ttarg%>"><%=UBBCode(adsrs("intro"))%></a>
<%          Case "gif"%><a href="<%=DqUrl%>/url.asp?id=<%=adsrs("id")%>" target="<%=ttarg%>"><img art="<%=adsrs("sitename")%>" border=0  <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>"></a> 
<%          Case "swf"%><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http:/download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"; <%=GaoAndKuan%>><param name=movie value="<%=adsrs("gif_url")%>"><param name=quality value=high>

  <embed src="<%=adsrs("gif_url")%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"  width="<%=adsrs("wid")%>" height="<%=adsrs("hei")%>"></embed></object>
<%           Case "dai"%><iframe marginwidth=0 marginheight=0  frameborder=0 bordercolor=000000 scrolling=no  name="忠网WEB广告管理系统 zon.cn" src="<%=DqUrl%>/daima.asp?id=<%=adsrs("id")%>"  <%=GaoAndKuan%>></iframe>

  <%          Case else%><a href="<%=DqUrl%>/url.asp?id=<%=adsrs("id")%>" target="<%=ttarg%>"><img art="<%=adsrs("sitename")%>"  border=0  <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>"></a>
<%
           End Select%>


<%
adsrs.close
%>

<%
    case "yulanggw"
    
%><script language="javascript" src="<%=DqUrl()%>/ad.asp?i=<%=getid%>"></script>
    
<%
    case "open"
set adsrs=server.createobject("adodb.recordset")
adssql="Select id,sitename,act from [ads] where id="&getid
adsrs.open adssql,adsconn,1,3
adsrs("act")=1
adsrs.Update
%>
<br><br>&#20197;&#19979;&#24191;&#21578;&#26465;&#34987;&#28608;&#27963;&#65306;<br><br>
<%=adsrs("sitename")%><br><br>
<%
adsrs.close

End Select
%>


<%
set adsrs=nothing 
adsconn.close
set adsconn=nothing
end if
%>
</td></tr><tr height=10 align=center><td><a href=javascript:window.close()>&#20851;&#38381;&#31383;&#21475;</a>
    <a href="javascript:this.location.reload();">&#21047;&#26032;</a></td></tr></table>
</center></div>
</body></html>