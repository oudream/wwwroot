<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body text="#000000" leftmargin="8" topmargin="8" marginwidth="0" marginheight="0">
  <!--#include file="dbm_adminconn.asp" -->
<%
dbm_hildden_id=request("dbm_hildden_id")
sql="select * from dbm_hildden where dbm_hildden_id=" & dbm_hildden_id 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
brief = replace(rs("brief"),chr(13),"<br>")
%>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
  
<TABLE width=600 
                  border=0 align="center" cellPadding=5 cellSpacing=1 bordercolor="#000000" class="tabp">
  <TBODY>
    <TR align="center" class="tdp"> 
      <TD colspan="2" class="tdp">��ַ�鿴 </TD>
    </TR>
    <TR class="tdp"> 
      <TD width="20%" align="right" class="tdp"> ���⣺</TD>
      <TD class="tdp"> <%=rs("htitle")%> </TD>
    </TR>
    <TR class="tdp"> 
      <TD align="right" class="tdp"> ��ַ��</TD>
      <TD class="tdp"> <%=rs("web_addr")%></TD>
    </TR>
    <TR class="tdp"> 
      <TD align="right" class="tdp">��飺</TD>
      <TD class="tdp"> <%=brief%> </TD>
    </TR>
  </TBODY>
</TABLE>

<%
rs.close
set rs=nothing
%>
</body>
</html>
