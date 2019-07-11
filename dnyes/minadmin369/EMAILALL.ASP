<!--#include file="../mymefaq5411jkjkh/favorible/showme.asp" -->
<!--#include file="../myjmail.asp" -->
<%

topic="ok thankyou"
mailbody="you are error,yu mast go go my herr"

Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select email from user order by id desc" 
rs.open sql,conn,1,1 

rs.pagesize=20
i=rs.pagecount
for j=1 to i 
rs.AbsolutePage=j
j=j+1
do while not rs.eof
if rs("email")<>"" then
email=rs("email")
call jmail(email,topic,mailbody)
end if
rs.movenext
n=n+1                                                                     
if n>=rs.pagesize then exit do                                                           
loop
next
rs.close
set rs=nothing
%> 
