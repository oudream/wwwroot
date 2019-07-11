<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; chaorset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" --> 
<%
if request("option")="edit" then
if request("tfconfirm")="yes" then sql="update orders set confirm='yes' where orders_id=" & request("orders_id")
if request("tfconfirm")="no" then sql="update orders set confirm='no' where orders_id=" & request("orders_id")
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Update order is complete. ');</SCRIPT>")
end if
%>
<%
if request("option")="del" then
conn.execute "delete * from orders where orders_id=" & request("orders_id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del the order is complete. ');</SCRIPT>")
end if
%>
<%
osql="select * from orders order by orders_id desc" 
Set ors=Server.CreateObject("ADODB.RecordSet") 
ors.Open osql,conn,1,1 

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

ors.pagesize=20
ors.AbsolutePage=pagecount
%> 
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td>&nbsp;</td>
    <td height="50">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td align="center" class="header">ALL ORDER</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td> <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
              <%
if ors.eof or ors.bof then
response.Write("<tr bgcolor=#FFFFFF><td height=25 colspan=11><b>NO DATA</b></td></tr></table>")
response.End()
ors.close
set ors=nothing
end if
%>
              <tr> 
                <td width="56" align="center" bgcolor="#FFFFFF"><b>uid</b></td>
                <td width="56" align="center" bgcolor="#FFFFFF"><b>name</b></td>
                <td width="89" align="center" bgcolor="#FFFFFF"><b>telphone</b></td>
                <td width="224" align="center" bgcolor="#FFFFFF"><b>product</b></td>
                <td width="224" align="center" bgcolor="#FFFFFF"><b>quantity</b></td>
                <td width="224" align="center" bgcolor="#FFFFFF"><b>datetime</b></td>
                <td width="224" align="center" bgcolor="#FFFFFF"><b>confirm</b></td>
                <td colspan="2" align="center" bgcolor="#FFFFFF"><strong>Operation</strong></td>
              </tr>
              <%
i=0
do while not ors.eof
%>
              <tr> 
                <td height="35" align="center" bgcolor="#f5f5f5"><a href="user_show.asp?page=<%=pagecount%>&pid=<%=ors("pid")%>"> 
                  <%
		  	psql="select * from player where pid=" & ors("pid") 
			Set prs=Server.CreateObject("ADODB.RecordSet") 
			prs.Open psql,conn,1,1 
			if not prs.eof then
			response.Write(prs("uid"))
			end if
			prs.close
		  %>
                  </a>&nbsp; </td>
                <td align="center" bgcolor="#f5f5f5"><%=ors("name")%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=ors("tel")%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"> <%
		  	psql="select * from subs where subs_id=" & ors("subs_id") 
			prs.Open psql,conn,1,1 
			if not prs.eof then
			response.Write(prs("name"))
			end if
			prs.close
			set prs=nothing
		  %> &nbsp; </td>
                <td align="center" bgcolor="#f5f5f5"><%=ors("count")%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=ors("inserttime")%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=ors("confirm")%>&nbsp;</td>
                <td width="224" align="center" bgcolor="#f5f5f5"> <a href="Shopping_Order_V.asp?page=<%=pagecount%>&orders_id=<%=ors("orders_id")%>&option=edit&tfconfirm=yes">YES 
                  </a> / <a href="Shopping_Order_V.asp?page=<%=pagecount%>&orders_id=<%=ors("orders_id")%>&option=edit&tfconfirm=no">NO</a> 
                </td>
                <td width="224" align="center" bgcolor="#f5f5f5"> <a href="Shopping_Order_Show.asp?page=<%=pagecount%>&orders_id=<%=ors("orders_id")%>">View</a> 
                  / <a href="Shopping_Order_V.asp?page=<%=pagecount%>&orders_id=<%=ors("orders_id")%>&option=del" onClick="return confirm('Are you sure to del it?');">Del</a> 
                </td>
              </tr>
              <%
ors.movenext
i=i+1                                                                     
if i>=ors.pagesize then exit do                                                           
loop
%>
            </table>
            <!-- ========================================================================================================		  

													fixture or result ============stop

 ======================================================================================================== -->
          </td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td> <table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
for zi=1 to ors.pagecount
%>
                <td width="50"><a href="Shopping_Order_V.asp?page=<%=zj%>">|&nbsp;<%=zj%>&nbsp;|</a></td>
                <%
zj=zj+1
next
%>
              </tr>
            </table></td>
          <td width="20%" align="left"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td> <div align="center"> Record :<font color=red><b><%=ors.recordcount%></b></font> Page <font color=red><b><%=pagecount%></b></font> of <font color=red><b><%=ors.pagecount%></b></font> </div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="50" colspan="3">&nbsp;</td>
  </tr>
</table>
<%
ors.close
set ors=nothing
%>
