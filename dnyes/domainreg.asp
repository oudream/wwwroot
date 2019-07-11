<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%
Session("domain")=replace(request("domain"),"'","¡¯")
subsname=replace(request("subsname"),"'","¡¯")
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
ProductList = Session("ProductList")
Products = Split(add, ", ")
For I=0 To UBound(Products)
   PutToShopBag Products(I), ProductList
Next
Session("ProductList") = ProductList
conn.close
set conn=nothing
response.redirect"check.asp"
%>