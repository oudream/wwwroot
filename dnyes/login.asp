<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="mdf.asp" -->
<%
if request("logout")="yes" then
            session("y_it_uid")=""
            session("y_it_pwd")=""
            session("y_it_userdiscount")=""
            session("y_it_userlevel")=""
            session("y_it_prepay")=""
            Session("ProductList")=""
            Session("subsList")=""
            session("Quatityt")=""
			Session("ProductList")=""
			Session("domain")=""
			session("sum")=""
			session("gyslist")=""
			session("ordernote")=""
			session("price")=""
			session("subid")=""
url="index.asp"
response.redirect url
response.end
end if
%>
<%
uid=replace(request("uid"),"'","¡¯")
pwd=replace(request("pwd"),"'","¡¯")
if uid="" or pwd="" then
response.redirect "error.asp?err=002"
response.end
end if

pwd=md5(pwd)
sql="select * from user where uid='"&uid&"'"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
if rs.eof or rs.bof then
response.redirect "error.asp?err=002"
response.end
end if

if pwd<>rs("pwd") then
response.redirect "error.asp?err=002"
response.end
end if

userlevel=rs("userlevel")
session("y_it_uid")=rs("uid")
session("y_it_pwd")=rs("pwd")
session("y_it_prepay")=rs("prepay")
session("y_it_userlevel")=userlevel
session("y_it_usermail")=trim(rs("email"))
rs.close
set rs=nothing


sql="select * from userlevel where userlevel='"&userlevel&"'"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
discount=rs("discount")/100
discount=1-discount
session("y_it_userdiscount")=discount
rs.close
set rs=nothing

url=request("url")
response.redirect url
'response.write session("y_it_uid") & "<br>" & session("y_it_pwd") & "<br>" & session("y_it_userdiscount")
'response.end
%>
