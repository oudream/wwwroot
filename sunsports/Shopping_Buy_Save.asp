<!--#include file="conn/conn.asp" -->
<%
if request("option")="add" then
user_pid=session("user_pid")
subs_id=request("subs_id")
zsize=request("size")
zsize=replace(zsize,"'","''")
zcount=request("count")
zpayment=request("payment")
zpayout=request("payout")
zcontent=request("Content")
zcontent=replace(zcontent,"'","''")
zuname=request("uname")
zuname=replace(zuname,"'","''")
ztel=request("tel")
ztel=replace(ztel,"'","''")
zip=request("zip")
zip=replace(zip,"'","''")
zcontact=request("contact")
zcontact=replace(zcontact,"'","''")
zemail=request("email")
zemail=replace(zemail,"'","''")
if zsize="" then zsize=" "
if zcontent="" then zcontent=" "
if zemail="" then zemail=" "
if user_pid<>"" then 
sql="insert into orders (pid,subs_id,count,size,payment_id,payout_id,content,name,email,tel,zip,contact,inserttime) values ("&user_pid&","&subs_id&","&zcount&",'"&zsize&"',"&zpayment&","&zpayout&",'"&zcontent&"','"&zuname&"','"&zemail&"','"&ztel&"','"&zip&"','"&zcontact&"','"&now()&"')"
else
sql="insert into orders (subs_id,count,size,payment_id,payout_id,content,name,email,tel,zip,contact,inserttime) values ("&subs_id&","&zcount&",'"&zsize&"',"&zpayment&","&zpayout&",'"&zcontent&"','"&zuname&"','"&zemail&"','"&ztel&"','"&zip&"','"&zcontact&"','"&now()&"')"
end if
conn.Execute(sql)
response.Redirect("Shopping_Buy_Check.asp")
end if
%>
<%
if request("option")="edit" then
user_pid=session("user_pid")
subs_id=request("subs_id")
zsize=request("size")
zsize=replace(zsize,"'","''")
zcount=request("count")
zpayment=request("payment")
zpayout=request("payout")
zcontent=request("Content")
zcontent=replace(zcontent,"'","''")
zuname=request("uname")
zuname=replace(zuname,"'","''")
ztel=request("tel")
ztel=replace(ztel,"'","''")
zip=request("zip")
zip=replace(zip,"'","''")
zcontact=request("contact")
zcontact=replace(zcontact,"'","''")
zemail=request("email")
zemail=replace(zemail,"'","''")
if zsize="" then zsize=" "
if zcontent="" then zcontent=" "
if zemail="" then zemail=" "
if user_pid<>"" then 
sql="update orders set pid="&user_pid&",subs_id="&subs_id&",count="&zcount&",size='"&zsize&"',payment_id="&zpayment&",payout_id="&zpayout&",content='"&zcontent&"',name='"&zuname&"',email='"&zemail&"',tel='"&ztel&"',zip='"&zip&"',contact='"&zcontact&"' where orders_id=" & request("orders_id")
else
sql="update orders set subs_id="&subs_id&",count="&zcount&",size='"&zsize&"',payment_id="&zpayment&",payout_id="&zpayout&",content='"&zcontent&"',name='"&zuname&"',email='"&zemail&"',tel='"&ztel&"',zip='"&zip&"',contact='"&zcontact&"' where orders_id=" & request("orders_id")
end if
conn.Execute(sql)
response.Redirect("Shopping_Buy_Check.asp?orders_id="&request("orders_id"))
end if
%>
