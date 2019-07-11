<!--#include file=config.asp --><html><head><title>忠网广告 业务电话:010-82771022</title></head><body topmargin="0" leftmargin="0">
<%
Dim GaoAndKuan
GaoAndKuan=""
adsconn.open adsdata
set adsrs=server.createobject("adodb.recordset")
adssql="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei from [ads] where id="&trim(request.querystring("i"))
adsrs.open adssql,adsconn,3,3
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


           Select Case adsrs("xslei")
    
            Case "txt"%><a title="<%=adsrs("intro")%>" href="<%=DqUrl%>/url.asp?id=<%=adsrs("id")%>" target="<%=ttarg%>"><font color=red><%=adsrs("sitename")%></font></a>
<%          Case "gif"%><a title="<%=adsrs("intro")%>" href="<%=DqUrl%>/url.asp?id=<%=adsrs("id")%>" target="<%=ttarg%>"><img border=0  <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>"></a> 
<%          Case "swf"%><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http:/download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"; <%=GaoAndKuan%>><param name=movie value="<%=adsrs("gif_url")%>"><param name=quality value=high>
  <%          Case "dai"%><%=adsrs("intro")%>
  <embed src="<%=adsrs("gif_url")%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"></embed></object>
<%          Case else%><a title="<%=adsrs("intro")%>" href="<%=DqUrl%>/url.asp?id=<%=adsrs("id")%>" target="<%=ttarg%>"><img border=0  <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>"></a>
<%
           End Select%><%
adsrs.close
set adsrs=nothing
adsconn.close
set adsconn=nothing 
%></body></html>