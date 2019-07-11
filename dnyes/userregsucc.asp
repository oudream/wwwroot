<!--#include file="mymefaq5411jkjkh/favorible/showme.asp" -->
<!--#include file="myjmail.asp" -->
<%
dim zadduser, zuid, zpwd, zcom_cn, zcom_en, zcontact_cn, zcontact_en, zu_or_i, ztechcc, ztechsp, ztechsp_cn, ztechen, ztechen_cn, zcityz, zcitye, zaddr, zaddr_en, zpostnumber, ztel1, ztel2, zfax, zqq, zemail, zmemo
zuid=trim(request("uid"))		'必要的 
zpwd=trim(request("pwd"))		'必要的  
zpwdb=trim(request("pwdb"))		'必要的  
zcom_cn=request("com_cn")
zcom_en=request("com_en")
zcontact_cn=request("contact_cn")		'必要的 
zcontact_en=request("contact_en")
zu_or_i=request("u_or_i")		'必要的 
ztechcc=request("techcc")		'必要的 
ztechsp=request("techsp")
ztechsp_cn=request("techsp_cn")
ztechen=request("techen")
ztechen_cn=request("techen_cn")
zcityz=request("cityz")		'必要的 
zcitye=request("citye")
zaddr=request("addr")		'必要的 
zaddr_en=request("addr_en")
zpostnumber=request("postnumber")		'必要的 
ztel1=request("tel1")		'必要的 
ztel2=request("tel2")
zfax=request("fax")
zqq=request("qq")
zemail=request("email")		'必要的 
zmemo=request("memo")
zusertype=request("usertype")
ztime=now()
%>

<%
if trim(zcom_cn)="" then zcom_cn=" "
if trim(zcom_en)="" then zcom_en=" "
if trim(zcontact_cn)="" then zcontact_cn=" "
if trim(zcontact_en)="" then zcontact_en=" "
if trim(ztechsp)="" then ztechsp=" "
if trim(ztechsp_cn)="" then ztechsp_cn=" "
if trim(ztechen)="" then ztechen=" "
if trim(ztechen_cn)="" then ztechen_cn=" "
if trim(zcitye)="" then zcitye=" "
if trim(zaddr_en)="" then zaddr_en=" "
if trim(ztel2)="" then ztel2=" "
if trim(fax)="" then fax=" "
if trim(zqq)="" then zqq=" "
if trim(zmemo)="" then zmemo=" "

if trim(zusertype)<>"" then
	if zusertype="buss" then
	usertypestr="企业用户"
	else
	usertypestr="个人用户"
	end if
else
	usertypestr=" "
end if   
%>

<%
if trim(ztechcc)="CHINA" then 
shengz=ztechsp_cn
shenge=ztechen_cn
else
shengz=ztechsp
shenge=ztechen
end if
'dim内定义的
'zadduser, zuid, zpwd, zcom_cn, zcom_en, zcontact_cn, zcontact_en, zu_or_i, ztechcc, ztechsp, ztechsp_cn, ztechen, 
'ztechen_cn, zcityz, zcitye, zaddr, zaddr_en, zpostnumber, ztel1, ztel2, zfax, zqq, zemail, zmemo

'request接收的
'zuid zpwd zcom_cn zcom_en zcontact_cn zcontact_en zu_or_i ztechcc ztechsp ztechsp_cn ztechen 
'ztechen_cn zcityz zcitye zaddr zaddr_en zpostnumber ztel1 ztel2 zfax zqq zemail zmemo 

'数据库内user表的
'"uid" "pwd" "namez" "namee" "contactz" "contacte" "govz" "gove" "shengz" "shenge" "cityz" "citye" "oicq" "email" 
'"fax" "postnumber" "dizhiz" "dizhie" "age" "sex" "tel2" "tel" "userlevel" "prepay" "sumjifen" "memo"

'================================body/
'================================body\
%>


<%on error resume next  '不支持组件时忽略错误
'进行安全性监测，看数据来源是否是本服务器页面
if not instr(1,Request.ServerVariables("http_Referer"),Request.ServerVariables ("SERVER_NAME"),1)=8 then
response.write "<b>请不要从非本服务器的页面提交信息</b>"
'response.redirect "error.asp?err=001"
response.end
end if
%>



 <%
'开始向数据库写入数据，并检测是否已有此用户
set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM user where uid= '" & zuid & "'"
rs.open sql,conn,1,1

if not (rs.Bof or rs.eof) then
rs.Close
set rs=nothing
response.redirect "error.asp?err=004"
response.end
else
sql="insert into user(uid,pwd,email,sex,namez,namee,govz,gove,shengz,shenge,cityz,citye,dizhiz,dizhie,tel,postnumber,oicq,fax,tel2,contactz,contacte,regtime,memo,usertype) values ('"&zuid&"','"&zpwdb&"','"&zemail&"','"&zu_or_i&"','"&zcom_cn&"','"&zcom_en&"','"&ztechcc&"','"&ztechcc&"','"&shengz&"','"&shenge&"','"&zcityz&"','"&zcitye&"','"&zaddr&"','"&zaddr_en&"','"&ztel1&"','"&zpostnumber&"','"&zqq&"','"&zfax&"','"&ztel2&"','"&zcontact_cn&"','"&zcontact_en&"','"&ztime&"','"&zmemo&"','"&usertypestr&"')"
conn.Execute(sql)
'给用户发信，提示已经注册成功
%>
<%
topic="恭喜您在 - 信网 数信科技 - 注册成功"
call getemailbody1(emailbody,zuid,zpwd)
call Jmail2(zemail,topic,emailbody)
%>

<%
end if
%> 

<%
sql="select * from user where uid='"&zuid&"' order by id desc"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
if rs.eof or rs.bof then
response.redirect "error.asp?err=002"
response.end
end if

session("y_it_uid")=rs("uid")
session("y_it_pwd")=rs("pwd")
session("y_it_prepay")=rs("prepay")
session("y_it_userlevel")=rs("userlevel")
session("y_it_usermail")=trim(rs("email"))
session("y_it_userdiscount")=1
rs.close
set rs=nothing

Session("ProductList")=""
Session("subsList")=""
session("Quatityt")=""
Session("domain")=""
session("sum")=""
session("gyslist")=""
session("ordernote")=""
session("price")=""
session("subid")=""

%>
<%
domainurl=session("domainurl")
if trim(domainurl)<>"" then
response.Redirect(domainurl)
response.end()
end if
%>

<%
if session("y_it_uid")=zuid then
response.Redirect("userregsuccsucc.asp?pwd="&zpwd)
response.End()
end if
%>
