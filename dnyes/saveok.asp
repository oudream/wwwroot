<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="myjmail.asp" -->
<%
on error resume next  '不支持组件时忽略错误
if Session("ProductList")="" or session("y_it_uid")="" then 
response.redirect "error.asp?err=007"
response.end
end if
ifpay="true"
paymenttype=request("paymenttype")
'产生订单号（内部和外部）及定单产生日期及其他信息
'交易日期，格式：YYYYMMDD
yy=year(date)
mm=right("00"&month(date),2)
dd=right("00"&day(date),2)
riqi=yy & mm & dd

'生成订单号所有所需元素,格式为：小时，分钟，秒
xiaoshi=right("00"&hour(time),2)
fenzhong=right("00"&minute(time),2)
miao=right("00"&second(time),2)
shijian=xiaoshi & fenzhong & miao
'生成订单号
inBillNo=riqi & "-" & shijian

'进行安全性监测，看数据来源是否是本服务器页面
if not instr(1,Request.ServerVariables("http_Referer"),Request.ServerVariables ("SERVER_NAME"),1)=8 then
response.write "<b>请不要从非本服务器的页面提交信息</b>"
'response.redirect "error.asp?err=001"
response.end
end if
paymenttype=request("paymenttype")
'从提交表单返回值，并查看是否有非法自符
i=request("i")
for j=1 to i
matemp="glma"&j
fm=request(matemp)
if fm="" then  fm= " "
if j=1 then glma=fm
if j>1 then glma=glma&"," & fm & ""
zhtemp="glzh"&j
fh=request(zhtemp)
if fh="" then fh= " "
if j=1 then glzh=fh
if j>1 then glzh=glzh&"," & fh & ""
next


if session("ordernote")="" then session("ordernote")=" "

if session("subid")="" then session("subid")=" "

conn.execute "insert into orders (username,inBillNo,ordertime,summoney,paymenttype,glzh,glma,userlevel,xxgys,xiangxi,xiangxi1,ordernote,subid) values ('"&session("y_it_uid")&"','"&inBillNo&"','"&now()&"','"&session("sum")&"','"&paymenttype&"','"&glzh&"','"&glma&"','"&session("y_it_userlevel")&"','"&session("gyslist")&"','"&session("subslist")&"','"&session("Quatityt")&"','"&session("ordernote")&"','"&session("subid")&"')"


sqlp2 = "SELECT * FROM user where uid= '" &session("y_it_uid")& "'"
set rsp2 = Server.CreateObject("ADODB.Recordset")
rsp2.open sqlp2,conn,1,1
zcontactz=rsp2("contactz")
zemail=rsp2("email")
zsumjifen=rsp2("sumjifen")
rsp2.close
set rsp2=nothing

if paymenttype="预付款支付" then
session("y_it_prepay")=clng(session("y_it_prepay"))-clng(session("sum"))
zsumjifen=zsumjifen+clng(session("sum"))
conn.execute "update user set prepay='"&session("y_it_prepay")&"' where uid='"&session("y_it_uid")&"'"
conn.execute "update user set sumjifen='"&zsumjifen&"' where uid='"&session("y_it_uid")&"'"
end if


sqlp1="select * from paydefault where paymenttype='"&paymenttype&"'"
set rsp1=server.createobject("adodb.recordset") 
rsp1.open sqlp1,conn,1,1 
paymenttype=rsp1("paymenttype")
paymentmessage = rsp1("paymentmessage")
paymentmessage = replace(paymentmessage,chr(13),"<br>")
paymentmessage = replace(paymentmessage,chr(32),"&nbsp;")  
paymentmessage1 = rsp1("paymentmessage")
rsp1.close
set rsp1=nothing

%>

<%
' inBillNo
subid=session("subid")
subsname=session("subslist")
price=session("price")
amount=session("Quatityt")
gross=session("sum")
nowtime=now()
%>
<%
topic="信网 数信科技-订单提交通知信"
call getemailbody2(emailbody,inBillNo,subid,subsname,price,amount,gross,nowtime,paymenttype,paymentmessage)
call Jmail3(zemail,topic,emailbody)
%>

<%
saveok1 = inBillNo
saveok2 = session("sum")
Session("ProductList")=""
Session("subsList")=""
session("Quatityt")=""
Session("domain")=""
Session("domainurl")=""
session("sum")=""
session("gyslist")=""
session("ordernote")=""
session("subid")=""
session("price")=""

response.Redirect("saveokok.asp?paymenttype="&paymenttype&"&saveok1="&saveok1&"&saveok2="&saveok2)
response.End()
%>
