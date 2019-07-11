<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="mdf.asp" -->
<%on error resume next%>
<%
'进行安全性监测，看数据来源是否是本服务器页面
if not instr(1,Request.ServerVariables("http_Referer"),Request.ServerVariables ("SERVER_NAME"),1)=8 then
response.write "<b>请不要从非本服务器的页面提交信息</b>"
'response.redirect "error.asp?err=001"
response.end
end if
%>
<%
'从提交表单返回值，并查看是否有非法自符 
uid =request("uid")
pwd =request("pwd")
pwdold =request("pwdold")
pwdcon =request("pwdcon")
if trim(pwdold)<> trim(pwdcon) then 
response.Redirect("usermodpass11.asp?errorstr=yes")
response.End()
else
%>
<%
pwd=md5(pwd)
conn.execute "update user set pwd='"&pwd&"' where uid='"&uid&"'"
Session("ProductList")=""
session("Quatityt")=""
Session("domain")=""
Session("domainurl")=""
session("sum")=""
session("y_it_uid")=""
session("y_it_pwd")=""
session("y_it_userdiscount")=""
session("y_it_userlevel")=""
session("y_it_prepay")=""
Session("subsList")=""
end if
%>
<%
response.Redirect("usermodpassokok.asp")
response.End()
%>