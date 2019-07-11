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
    <td valign="top">
	
<table width="102%" border="0" align="center" cellpadding="0" cellspacing="1" style="margin-bottom: 6">
  <tr> 
    <td width="1%" height="25"><div align="center"><b></b><font color="#FFFFFF"><br>
        </font></div></td>
    <td width="99%"><b>Result: </b>&gt;</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="100" colspan="2"> <div align="center"></div>
      <table border="0" width="100%" cellspacing="0" cellpadding="0" height="5" align="center">
        <br>
        <%
sql="select * from subs where name <> ''"
hw_name=request("hw_name")
if hw_name<>"" then sql=sql & " and ( name like '%"&hw_name&"%' or code like '%"&hw_name&"%' or packaging like '%"&hw_name&"%' or madein like '%"&hw_name&"%' or content like '%"&hw_name&"%' or brand like '%"&hw_name&"%')"
sql=sql&" order by subs_id DESC"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,3,3
if rs.eof then
%>
        <tr> 
          <td width="1011">Sorry,no this kind of product . Please try another 
            key words .</td>
        </tr>
        <%
else

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=5
rs.AbsolutePage=pagecount

%>
        <tr> 
          <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <%
do while not rs.eof
%>
                    <tr> 
                      <td width="120" height="135" align="center"><a href="Shopping_ShowProduct.asp?subs_id=<%=rs("subs_id")%>" target="_blank"><img src="<%=rs("pic")%>" width="120" height="120" border="0"></a></td>
                <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                  </table>
                  <b>
				  <%=rs("name")%></b><br>
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
                    <tr valign="top"> 
                      <td height="35" colspan="2" align="center"><table width="70%" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td height="1" bgcolor="#6184A3"><img src="1.gif" width="1" height="1"></td>
                          </tr>
                        </table></td>
                    </tr>
                    <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
%>
                  </table></td>
        </tr>
        <tr> 
          <td>
      <br> <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td valign="middle"><div align="center"></div></td>
        </tr>
      </table>
		  </td>
        </tr>
        <tr>
                <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
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
                            <td width="20"><a href="Shopping_Search_Result.asp?hw_name=<%=hw_name%>&page=<%=zj%>"><%=zj%></a></td>
                            <%
zj=zj+1
next
%>
                          </tr>
                        </table></td>
                      <td width="150"> <table border="0" cellspacing="1" cellpadding="2" width="100%" align="center">
                          <tr> 
                            <td height="22"> Page : <b><font color=red><%=pagecount%></font>/<%=rs.pagecount%></b> 
                              <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%>
                              <a href="Shopping_Search_Result.ASP?hw_name=<%=hw_name%>&page=<%=cstr(pagecount+1)%>"> 
                              Next</a> 
                              <% end if %>
                              <% if rs.pagecount>1 and rs.pagecount=pagecount then %>
                              <a href="Shopping_Search_Result.ASP?hw_name=<%=hw_name%>&page=<%=cstr(pagecount-1)%>"> 
                              Prev</a> 
                              <%end if%>
                              <% if pagecount<>1 and rs.pagecount<>pagecount then%>
                              <a href="Shopping_Search_Result.ASP?hw_name=<%=hw_name%>&page=<%=cstr(pagecount-1)%>"> 
                              Prev</a> <a href="Shopping_Search_Result.ASP?hw_name=<%=hw_name%>&page=<%=cstr(pagecount+1)%>"> 
                              Next</a> 
                              <% end if%>
                            </td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
        </tr>

<%
end if
%>
      </table>
	  </td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="100" colspan="2"></td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
	
	
	</td>
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
<!--#include file="User_l.asp" -->
	 <br> <!--#include file="Shopping_Search.asp" -->
	 <br> 
      <!--#include file="Shopping_Hot_in.asp" -->
	 <br> 	
	</td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->