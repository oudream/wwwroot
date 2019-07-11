<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
	dim n,m,h,s
	dim y,z,show
	stats="所有头像"
	h=10
	s=50
	M=60
	call nav()
	call head_var(1)
	call main()
	call footer()
%>
<%sub main()%>
    <table cellpadding=6 cellspacing=1 class=tableborder1 align=center style="table-layout:fixed;word-break:break-all;">
    <tr>
    <th id=tabletitlelink>
<% if m>s then %>
<p align="center"><% for Y=1 to (m+s-1)/s %>
 [<a href=allface.asp?show=<%=Y%>> 头像列表<%=Y%> </a>]  
<% next 
end if%>
    </th>
    </tr>
<tr>
    <td class=tablebody1>
<table width="100%">
    <tr>
<% 
	dim p
	i=0
	p= request("show")
	if p="" then
		p=1
	end if
	z=s*p

	if z/m>1 then
		z=m
	else
		z=s*p
	end if

for n=s*p+1-s to z
     i=i+1
%>
      <td   align="center" height=50><IMG SRC=<%=Forum_info(7)%>Image<%=n%>.gif><br>Image<%=n%></td>
<%
	if i=h then
	response.write "</tr><tr>"
	i=0
	end if 
	next
%>

  </table>
    </td>
    </tr></table>
<%end sub%>