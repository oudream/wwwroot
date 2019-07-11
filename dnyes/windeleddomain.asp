<TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR> 
      <TD height="8"><img src="1.gif" width="1" height="1"></TD>
    </TR>
  </TBODY>
</TABLE>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<%
sqlnews="select * from deleddomain order by id desc" 
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
rsnews.Open sqlnews,conn,1,1 
n=0
do while not rsnews.eof
deleddatecon=rsnews("deleddate")
if deleddatecon=cstr(date) then 
	delcolor="red"
	elseif deleddatecon=cstr(DateAdd("d", -1, date)) then 
	delcolor="red"
	elseif deleddatecon=cstr(DateAdd("d", -2, date)) then 
	delcolor="red"
else
	delcolor=""
end if
%>
  <tr> 
    <td height="20" valign="top"><a href="showdeledtable.asp?id=<%=rsnews("id")%>" title="<%=rsnews("deledtable")%>" target="_blank"> 
     <font color=<%=delcolor%>><%=n+1%>¡¢&nbsp;<%=rsnews("deledtable")%>&nbsp;</font></a></td>
  </tr>
  <%
n=n+1
rsnews.movenext
if n>=5 then exit do
loop
rsnews.close
set rsnews=nothing
%>
  <tr> 
    <td align="right"><div align="right"><a href="viewdeleddomain.asp" target="_blank">¸ü¶à>>&nbsp;&nbsp;</a></div></td>
  </tr>
</table>
<TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR> 
      <TD height="9"><img src="1.gif" width="1" height="1"></TD>
    </TR>
  </TBODY>
</TABLE>
