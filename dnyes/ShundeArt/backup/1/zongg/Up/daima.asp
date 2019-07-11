<!--#include file=config.asp -->
<%  if trim(request("id"))<>"" then
    adsconn.open adsdata
    set adsrs=server.createobject("adodb.recordset")
    adssql="Select top 1 intro from [ads] where id="&trim(request("id"))&" order by time"
    adsrs.open adssql,adsconn,1,1       
    response.write adsrs(0)
    adsrs.close
    set adsrs=nothing
    set adsconn=nothing
    else
    response.write "<center><br><br>无效广告。</center>"
    end if
%>