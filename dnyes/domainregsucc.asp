<!--#include file="mymefaq5411jkjkh/favorible/showme.asp" -->
<%
zuid=trim(session("y_it_uid"))		'必要的 
zcom_cn=request("donc")
zcom_en=request("done")
zcontact_cn=request("doceoc")		'必要的 
zcontact_en=request("doceoe")
zshengz=request("dos")
zcityz=request("docc")		'必要的 
zcitye=request("doce")
zaddr=request("doac")		'必要的 
zaddr_en=request("doae")
zpostnumber=request("doz")		'必要的 
zgcom_cn=request("danc")
zgcom_en=request("dane")
zgcontact_cn=request("daoc")		'必要的 
zgcontact_en=request("daoe")
zgshengz=request("das")
zgcityz=request("dacc")		'必要的 
zgcitye=request("dace")
zgaddr=request("daac")		'必要的 
zgaddr_en=request("daae")
zgpostnumber=request("daz")		'必要的 

zteli=request("dapi")		'必要的 
ztela=request("dapa")		'必要的 
zteln=request("dapn")		'必要的 
ztele=request("dape")		'必要的 
zfaxi=request("dafi")
zfaxa=request("dafa")
zfaxn=request("dafn")
zfaxe=request("dafe")
zemail=request("daem")		'必要的 

if trim(ztele)<>"" then
ztel=trim(zteli)&"-"&trim(ztela)&"-"&trim(zteln)&"-"&trim(ztele)
else
ztel=trim(zteli)&"-"&trim(ztela)&"-"&trim(zteln)
end if

if trim(zfaxe)<>"" then
zfax=trim(zfaxi)&"-"&trim(zfaxa)&"-"&trim(zfaxn)&"-"&trim(zfaxe)
else
zfax=trim(zfaxi)&"-"&trim(zfaxa)&"-"&trim(zfaxn)
end if

ztime=now()

zd4=request("d4")
zd5=request("d5")
zd1=request("d1")
zd2=request("d2")

if trim(zd4)<>"" or trim(zd5)<>"" then
zdns1=zd4
zdns2=zd5
else
zdns1=zd1
zdns2=zd2
end if

if zdns1="" then zdns1=" "
if zdns2="" then zdns2=" "
%>


<%
'on error resume next  '不支持组件时忽略错误
'进行安全性监测，看数据来源是否是本服务器页面
if not instr(1,Request.ServerVariables("http_Referer"),Request.ServerVariables ("SERVER_NAME"),1)=8 then
response.write "<b>请不要从非本服务器的页面提交信息</b>"
'response.redirect "error.asp?err=001"
response.end
end if
%>

<%
sql="insert into domainregs(uid,namez,namee,contactz,contacte,shengz,cityz,citye,dizhiz,dizhie,postnumber,gnamez,gnamee,gcontactz,gcontacte,gshengz,gcityz,gcitye,gdizhiz,gdizhie,gpostnumber,gtel,gfax,gemail,dns1,dns2,regtime) values ('"&zuid&"','"&zcom_cn&"','"&zcom_en&"','"&zcontact_cn&"','"&zcontact_en&"','"&zshengz&"','"&zcityz&"','"&zcitye&"','"&zaddr&"','"&zaddr_en&"','"&zpostnumber&"','"&zgcom_cn&"','"&zgcom_en&"','"&zgcontact_cn&"','"&zgcontact_en&"','"&zgshengz&"','"&zgcityz&"','"&zgcitye&"','"&zgaddr&"','"&zgaddr_en&"','"&zgpostnumber&"','"&ztel&"','"&zfax&"','"&zemail&"','"&zdns1&"','"&zdns2&"','"&ztime&"')"
conn.Execute(sql)
%>
<%
Session("domain")=replace(request("domain"),"'","’")
if trim(request("after"))="" then 
subsname="通用网址"
else
subsname=replace(request("after"),"'","’")
end if
sql="select * from subs where subsname='"&subsname&"' order by id desc"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1
add=rs("bookbm")
rs.close
set rs=nothing

Sub PutToShopBag( add, ProductList )
   If Len(ProductList) = 0 Then
      ProductList = "'" & add & "'"
ElseIf InStr( ProductList, add ) <= 0 Then
      ProductList = ProductList & ", '" & add & "'"
   End If
End Sub
%>
<%
Products = Split(add, ", ")
For I=0 To UBound(Products)
   PutToShopBag Products(I), ProductList
Next
Session("ProductList") = ProductList
conn.close
set conn=nothing
response.redirect"check.asp"
%>