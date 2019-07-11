<!--#include file="Top_sun.asp"-->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181" valign="top"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="15"><img src="1.gif" width="1" height="1"></td>
        </tr>
      </table>
      <!--#include file="Shopping_Sort.asp" -->	</td>
    <td width="20">&nbsp; </td>
    <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="24">&nbsp;</td>
              </tr>
            </table>
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <%
sort_id=request("sort_id")
if sort_id="" then sort_id=49
sql="select * from subs where sort_id="& sort_id &" order by subs_id desc "
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' There are not products in the sort, please choose other sort ');location='Shopping_Sort_view.asp?sort_id=55';</SCRIPT>" )
response.end()
end if
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=5
rs.AbsolutePage=pagecount

do while not rs.eof
%>
              <tr> 
                <td width="120" height="135" align="center" valign="top"><a href="Shopping_ShowProduct.asp?subs_id=<%=rs("subs_id")%>" target="_blank"><img src="<%=rs("pic")%>" alt="Click &amp; View" width="120" height="120" border="0"></a></td>
                <td width="10">&nbsp;</td>
                <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                  </table>
                  <b><%=rs("name")%></b><br>
                  Price:<font color="#FF0000"> $ <%=rs("price")%></font> <br>
                  Description: 
                  <%
				if len(rs("content")) > 100 then
				response.Write(left(rs("content"),98)&"..")
				else
				response.Write(rs("content"))
				end if
				%>
                  <br>
                  <br> <a href="Shopping_Buy.asp?subs_id=<%=rs("subs_id")%>">[ ORDER NOW ]</a> 
                </td>
              </tr>
              <tr> 
                <td height="35" colspan="3" align="center" valign="top"><table width="70%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td height="1" bgcolor="#6184A3"><img src="1.gif" width="1" height="1"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
%>
            </table></td>
        </tr>
      </table> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="79"> TOTAL : <%=rs.recordcount%></td>
          <td><table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
zk=5
if rs.pagecount < 6 then zk=rs.pagecount
for zi=1 to zk
%>
                <td width="20"><a href="Shopping_Sort_View.asp?sort_id=<%=sort_id%>&page=<%=zj%>"><%=zj%></a></td>
                <%
zj=zj+1
next
%>
              </tr>
            </table></td>
          <td width="150">
<table border="0" cellspacing="1" cellpadding="2" width="100%" align="center">
  <tr> 
    <td height="22"> 
       Page : <b><font color=red><%=pagecount%></font>/<%=rs.pagecount%></b>  
        <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%>
                  <a href="Shopping_Sort_View.ASP?sort_id=<%=sort_id%>&page=<%=cstr(pagecount+1)%>"> 
                  Next</a> 
                  <% end if %> <% if rs.pagecount>1 and rs.pagecount=pagecount then %>
                  <a href="Shopping_Sort_View.ASP?sort_id=<%=sort_id%>&page=<%=cstr(pagecount-1)%>"> 
                  Prev</a> 
                  <%end if%> <% if pagecount<>1 and rs.pagecount<>pagecount then%>
                  <a href="Shopping_Sort_View.ASP?sort_id=<%=sort_id%>&page=<%=cstr(pagecount-1)%>"> 
                  Prev</a> <a href="Shopping_Sort_View.ASP?sort_id=<%=sort_id%>&page=<%=cstr(pagecount+1)%>"> 
                  Next</a> 
                  <% end if%>
          </td>
        </tr>
      </table>

	  </td>
        </tr>
      </table></td>
    <td width="20" valign="top"><table width="95%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="24">&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
          <td height="830" background="images/10x1_blue.gif">&nbsp;</td>
        </tr>
      </table></td>
    <td width="181" valign="top">
<%
rs.close
set rs=nothing
%>
<!--#include file="User_l.asp" -->
	 <br> <!--#include file="Shopping_Search.asp" -->
	 <br> <!--#include file="Shopping_Hot_in.asp" -->
	 <br> 	
	</td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->