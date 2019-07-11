<!--#include file="conn/conn.asp" -->
<%
if request("option")="add" then
uid=trim(request("uid"))
pwd=trim(request("pwd"))
if uid<>"" and pwd<>"" then
usql="select * from player where uid='"&uid&"'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' The Userid have existed , please enter other Userid . ');history.go( -1 );</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing
adminlevel=4
ptype="s"
uname=trim(request("uname"))
tel=trim(request("tel"))
zip=trim(request("zip"))
contact=trim(request("contact"))
email=trim(request("email"))
if uname="" then uname=" "
if tel="" then tel=" "
if zip="" then zip=" "
if contact="" then contact=" "
if email="" then email=" "
sql="insert into player (name,tel,zip,contact,email,uid,pwd,adminlevel,inserttime) values ('"&uname&"','"&tel&"','"&zip&"','"&contact&"','"&email&"','"&uid&"','"&pwd&"',"&adminlevel&",'"&now()&"')"
conn.Execute(sql)
response.Redirect("Shopping_Reg_Succ.asp?uid="&uid)
end if
end if
end if
%>

<%
if request("option")="edit" then
uname=trim(request("uname"))
tel=trim(request("tel"))
zip=trim(request("zip"))
contact=trim(request("contact"))
email=trim(request("email"))
if uname="" then uname=" "
if tel="" then tel=" "
if zip="" then zip=" "
if contact="" then contact=" "
if email="" then email=" "
sql="update player set name='"&uname&"',tel='"&tel&"',zip='"&zip&"',contact='"&contact&"',email='"&email&"' where pid=" & session("user_pid")
conn.Execute(sql)
response.Redirect("Shopping_Reg_Esucc.asp")
end if
%>

<%
if request("option")="editpwd" then
oldpwd=trim(request("oldpwd"))
if oldpwd<>session("user_pwd") then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' The Old Password is not a valid Password , please enter again . ');history.go( -1 );</SCRIPT>")
else
pwd=trim(request("pwd"))
sql="update player set pwd='"&pwd&"' where pid=" & session("user_pid")
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit password is complete . ');location='index.asp';</SCRIPT>")
end if
end if
%>
