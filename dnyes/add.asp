<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%On Error Resume Next%>
<%
if session("y_it_uid")="" then 
response.redirect "error.asp?err=006&add="&request("add")
response.end
end if
%>
<%
Sub PutToShopBag( add, ProductList )
      ProductList = "'" & add & "'"
End Sub
%>
<%
session("domain")=""
ProductList = Session("ProductList")
Products = Split(Request("add"), ", ")
For I=0 To UBound(Products)
   PutToShopBag Products(I), ProductList
Next
Session("ProductList") = ProductList
conn.close
set conn=nothing
response.redirect"check.asp"
%>