<!--#include file="mymefaq5411jkjkh/favorible/showme.asp" -->
<!--#include file="myjmail.asp" -->
<%
dim zadduser, zuid, zpwd, zcom_cn, zcom_en, zcontact_cn, zcontact_en, zu_or_i, ztechcc, ztechsp, ztechsp_cn, ztechen, ztechen_cn, zcityz, zcitye, zaddr, zaddr_en, zpostnumber, ztel1, ztel2, zfax, zqq, zemail, zmemo
zuid=trim(request("uid"))		'��Ҫ�� 
zpwd=trim(request("pwd"))		'��Ҫ��  
zpwdb=trim(request("pwdb"))		'��Ҫ��  
zcom_cn=request("com_cn")
zcom_en=request("com_en")
zcontact_cn=request("contact_cn")		'��Ҫ�� 
zcontact_en=request("contact_en")
zu_or_i=request("u_or_i")		'��Ҫ�� 
ztechcc=request("techcc")		'��Ҫ�� 
ztechsp=request("techsp")
ztechsp_cn=request("techsp_cn")
ztechen=request("techen")
ztechen_cn=request("techen_cn")
zcityz=request("cityz")		'��Ҫ�� 
zcitye=request("citye")
zaddr=request("addr")		'��Ҫ�� 
zaddr_en=request("addr_en")
zpostnumber=request("postnumber")		'��Ҫ�� 
ztel1=request("tel1")		'��Ҫ�� 
ztel2=request("tel2")
zfax=request("fax")
zqq=request("qq")
zemail=request("email")		'��Ҫ�� 
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
	usertypestr="��ҵ�û�"
	else
	usertypestr="�����û�"
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
'dim�ڶ����
'zadduser, zuid, zpwd, zcom_cn, zcom_en, zcontact_cn, zcontact_en, zu_or_i, ztechcc, ztechsp, ztechsp_cn, ztechen, 
'ztechen_cn, zcityz, zcitye, zaddr, zaddr_en, zpostnumber, ztel1, ztel2, zfax, zqq, zemail, zmemo

'request���յ�
'zuid zpwd zcom_cn zcom_en zcontact_cn zcontact_en zu_or_i ztechcc ztechsp ztechsp_cn ztechen 
'ztechen_cn zcityz zcitye zaddr zaddr_en zpostnumber ztel1 ztel2 zfax zqq zemail zmemo 

'���ݿ���user���
'"uid" "pwd" "namez" "namee" "contactz" "contacte" "govz" "gove" "shengz" "shenge" "cityz" "citye" "oicq" "email" 
'"fax" "postnumber" "dizhiz" "dizhie" "age" "sex" "tel2" "tel" "userlevel" "prepay" "sumjifen" "memo"

'================================body/
'================================body\
%>


<%on error resume next  '��֧�����ʱ���Դ���
'���а�ȫ�Լ�⣬��������Դ�Ƿ��Ǳ�������ҳ��
if not instr(1,Request.ServerVariables("http_Referer"),Request.ServerVariables ("SERVER_NAME"),1)=8 then
response.write "<b>�벻Ҫ�ӷǱ���������ҳ���ύ��Ϣ</b>"
'response.redirect "error.asp?err=001"
response.end
end if
%>



 <%
'��ʼ�����ݿ�д�����ݣ�������Ƿ����д��û�
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
'���û����ţ���ʾ�Ѿ�ע��ɹ�
%>
<%
topic="��ϲ���� - ���� ���ſƼ� - ע��ɹ�"
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
