<!--#include file=config.asp -->

<%
dim getid,getclick,geturl,adsrs,adssql,adsrsc,adssqlc,getip
adsconn.open adsdata
getid=cint(request.querystring("id"))
set adsrs=server.createobject("adodb.recordset")
adssql="Select id,url,click from [ads] where id="&getid
adsrs.open adssql,adsconn,1,3
getclick=adsrs("click")+1
adsrs("click")=getclick
adsrs.Update

getip=request.ServerVariables("REMOTE_ADDR")
set adsrsc=server.createobject("adodb.recordset")
adssqlc="select * from iplist"
adsrsc.open adssqlc,adsconn,1,3
adsrsc.AddNew
adsrsc("adid") =getid
adsrsc("time") = now()
adsrsc("ip") = getip
adsrsc("class") = 2
adsrsc.update
adsrsc.close
set adsrsc=nothing 


geturl=adsrs("url")
adsrs.close
set adsrs=nothing 
adsconn.close
set adsconn=nothing
response.redirect geturl
%>