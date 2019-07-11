<!--#include file="conn/conn.asp" -->
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
          <td> <table width="100%" border="1">
              <%
do while not rs.eof
%>
              <tr> 
                <td width="39%" align="center"><a href="Shopping_ShowProduct.asp?subs_id=<%=rs("subs_id")%>" target="_blank"><img src="<%=rs("pic")%>" width="90" height="120"></a></td>
                <td width="61%">Name: <%=rs("name")%><br> <%
				if len(rs("content")) > 100 then
				response.Write(left(rs("content"),98)&"..")
				else
				response.Write(rs("content"))
				end if
				%> <br>
                  Price:<font color="#FF0000"> $ <%=rs("price")%></font> <br> <a href="Shopping_Buy.asp?subs_id=<%=rs("subs_id")%>">&lt;BUY&gt;</a> 
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
        <tr> 
          <td>
      <br> <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td valign="middle"><div align="center"><img src="IMAGES/pics/seperator.gif" width="359" height="1"></div></td>
        </tr>
      </table>
		  </td>
        </tr>
        <tr>
          <td><table width="100%" border="1">
        <tr> 
          <td width="25%"> TOTAL : <%=rs.recordcount%></td>
          <td width="32%"><table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
zk=5
if rs.pagecount < 6 then zk=rs.pagecount
for zi=1 to zk
%>
                <td width="50"><a href="Shopping_Search_Result.asp?hw_name=<%=hw_name%>&page=<%=zj%>"><%=zj%></a></td>
                <%
zj=zj+1
next
%>
              </tr>
            </table></td>
          <td width="43%"> <table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="65%" align="center">
              <tr bgcolor="#FFCC33"> 
                <td height="22"> <div align="center">page : <b><font color=red><%=pagecount%></font>/<%=rs.pagecount%></b> 
                    <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%>
                    <a href="Shopping_Search_Result.asp?hw_name=<%=hw_name%>&page=<%=cstr(pagecount+1)%>"> 
                    <font color="#000000">next</font></a> 
                    <% end if %>
                    <% if rs.pagecount>1 and rs.pagecount=pagecount then %>
                    <a href="Shopping_Search_Result.asp?hw_name=<%=hw_name%>&page=<%=cstr(pagecount-1)%>"> 
                    <font color="#000000">prev</font></a> 
                    <%end if%>
                    <% if pagecount<>1 and rs.pagecount<>pagecount then%>
                    <a href="Shopping_Search_Result.asp?hw_name=<%=hw_name%>&page=<%=cstr(pagecount-1)%>"> 
                    <font color="#000000">prev</font></a> <a href="Shopping_Search_Result.asp?hw_name=<%=hw_name%>&page=<%=cstr(pagecount+1)%>"> 
                    <font color="#000000">next</font></a> 
                    <% end if%>
                  </div></td>
              </tr>
            </table></td>
        </tr>
      </table>
	  </td>
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
