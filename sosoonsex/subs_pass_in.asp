<!--#include file="adminconn.asp" -->
<%
if request("option")="add" then

subs_pass_type=session("subs_pass_type")
gid=request("gid")
gcontacter=request("gcontacter")
if gcontacter="" then gcontacter=" "
pyear=request("pyear")
pmonth=request("pmonth")
pday=request("pday")
operationer_id=request("operationer_id")
if operationer_id="" then operationer_id=0
user_id=session("user_id")

if session("subs_id_1")<>0 then
mount_1=request("mount_1")
price_1=request("price_1")
counter_1=request("counter_1")
memo_1=request("memo_1")
if memo_1="" then memo_1=" "
sql="insert into subspass (subspass_type,gid,gcontacter,pyear,pmonth,pday,subs_id,mount,price,counter,memo,operationer_id,user_id,inserttime) values ('"&session("subs_pass_type")&"',"&gid&",'"&gcontacter&"',"&pyear&","&pmonth&","&pday&","&session("subs_id_1")&","&mount_1&","&price_1&","&counter_1&",'"&memo_1&"',"&operationer_id&","&session("user_id")&",'"&now()&"')"
conn.Execute(sql)
end if

if session("subs_id_2")<>0 then
mount_2=request("mount_2")
price_2=request("price_2")
counter_2=request("counter_2")
memo_2=request("memo_2")
if memo_2="" then memo_2=" "
sql="insert into subspass (subspass_type,gid,gcontacter,pyear,pmonth,pday,subs_id,mount,price,counter,memo,operationer_id,user_id,inserttime) values ('"&session("subs_pass_type")&"',"&gid&",'"&gcontacter&"',"&pyear&","&pmonth&","&pday&","&session("subs_id_2")&","&mount_2&","&price_2&","&counter_2&",'"&memo_2&"',"&operationer_id&","&session("user_id")&",'"&now()&"')"
conn.Execute(sql)
end if

if session("subs_id_3")<>0 then
mount_3=request("mount_3")
price_3=request("price_3")
counter_3=request("counter_3")
memo_3=request("memo_3")
if memo_3="" then memo_3=" "
sql="insert into subspass (subspass_type,gid,gcontacter,pyear,pmonth,pday,subs_id,mount,price,counter,memo,operationer_id,user_id,inserttime) values ('"&session("subs_pass_type")&"',"&gid&",'"&gcontacter&"',"&pyear&","&pmonth&","&pday&","&session("subs_id_3")&","&mount_3&","&price_3&","&counter_3&",'"&memo_3&"',"&operationer_id&","&session("user_id")&",'"&now()&"')"
conn.Execute(sql)
end if

if session("subs_id_4")<>0 then
mount_4=request("mount_4")
price_4=request("price_4")
counter_4=request("counter_4")
memo_4=request("memo_4")
if memo_4="" then memo_4=" "
sql="insert into subspass (subspass_type,gid,gcontacter,pyear,pmonth,pday,subs_id,mount,price,counter,memo,operationer_id,user_id,inserttime) values ('"&session("subs_pass_type")&"',"&gid&",'"&gcontacter&"',"&pyear&","&pmonth&","&pday&","&session("subs_id_4")&","&mount_4&","&price_4&","&counter_4&",'"&memo_4&"',"&operationer_id&","&session("user_id")&",'"&now()&"')"
conn.Execute(sql)
end if

if session("subs_id_5")<>0 then
mount_5=request("mount_5")
price_5=request("price_5")
counter_5=request("counter_5")
memo_5=request("memo_5")
if memo_5="" then memo_5=" "
sql="insert into subspass (subspass_type,gid,gcontacter,pyear,pmonth,pday,subs_id,mount,price,counter,memo,operationer_id,user_id,inserttime) values ('"&session("subs_pass_type")&"',"&gid&",'"&gcontacter&"',"&pyear&","&pmonth&","&pday&","&session("subs_id_5")&","&mount_5&","&price_5&","&counter_5&",'"&memo_5&"',"&operationer_id&","&session("user_id")&",'"&now()&"')"
conn.Execute(sql)
end if

if session("subs_id_6")<>0 then
mount_6=request("mount_6")
price_6=request("price_6")
counter_6=request("counter_6")
memo_6=request("memo_6")
if memo_6="" then memo_6=" "
sql="insert into subspass (subspass_type,gid,gcontacter,pyear,pmonth,pday,subs_id,mount,price,counter,memo,operationer_id,user_id,inserttime) values ('"&session("subs_pass_type")&"',"&gid&",'"&gcontacter&"',"&pyear&","&pmonth&","&pday&","&session("subs_id_6")&","&mount_6&","&price_6&","&counter_6&",'"&memo_6&"',"&operationer_id&","&session("user_id")&",'"&now()&"')"
conn.Execute(sql)
end if

if session("subs_id_7")<>0 then
mount_7=request("mount_7")
price_7=request("price_7")
counter_7=request("counter_7")
memo_7=request("memo_7")
if memo_7="" then memo_7=" "
sql="insert into subspass (subspass_type,gid,gcontacter,pyear,pmonth,pday,subs_id,mount,price,counter,memo,operationer_id,user_id,inserttime) values ('"&session("subs_pass_type")&"',"&gid&",'"&gcontacter&"',"&pyear&","&pmonth&","&pday&","&session("subs_id_7")&","&mount_7&","&price_7&","&counter_7&",'"&memo_7&"',"&operationer_id&","&session("user_id")&",'"&now()&"')"
conn.Execute(sql)
end if

if session("subs_id_8")<>0 then
mount_8=request("mount_8")
price_8=request("price_8")
counter_8=request("counter_8")
memo_8=request("memo_8")
if memo_8="" then memo_8=" "
sql="insert into subspass (subspass_type,gid,gcontacter,pyear,pmonth,pday,subs_id,mount,price,counter,memo,operationer_id,user_id,inserttime) values ('"&session("subs_pass_type")&"',"&gid&",'"&gcontacter&"',"&pyear&","&pmonth&","&pday&","&session("subs_id_8")&","&mount_8&","&price_8&","&counter_8&",'"&memo_8&"',"&operationer_id&","&session("user_id")&",'"&now()&"')"
conn.Execute(sql)
end if

if session("subs_id_1")=0 and session("subs_id_2")=0 and session("subs_id_3")=0 and session("subs_id_4")=0 and session("subs_id_5")=0 and session("subs_id_6")=0 and session("subs_id_7")=0 and session("subs_id_8")=0 then
response.Redirect("subs_select.asp?subs_id_all=no&passtype="&session("subs_pass_type"))
end if

response.Redirect("subs_select.asp?subs_id_all=yes&passtype="&session("subs_pass_type"))
end if
%>