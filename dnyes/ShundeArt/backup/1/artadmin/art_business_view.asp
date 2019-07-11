<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>

<!--#include file="adminconn.asp" -->
<%
sql="select * from art_business order by allname" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' None data ! ');</SCRIPT>")
response.End()
end if
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=20
rs.AbsolutePage=pagecount
%>
<table width="97%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr bgcolor="#FFFFFF"> 
    <td height="22" colspan="8" align="center" class="header"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="75%"> <table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
for zi=1 to rs.pagecount
%>
                <td width="50"><a href="art_business_view.asp?page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
                <%
zj=zj+1
next
%>
              </tr>
            </table></td>
          <td width="26%" align="left"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td align="right"> Record :<font color=red><b><%=rs.recordcount%></b></font> Page <font color=red><b><%=pagecount%></b></font> of <font color=red><b><%=rs.pagecount%></b></font> </td>
              </tr>
            </table></td>
        </tr>
      </table> </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="22" colspan="8" align="center" class="header">客户列表</td>
  </tr>
  <tr> 
    <td width="80" height="30" bgcolor="#FFFFFF">全称</td>
    <td width="60" bgcolor="#FFFFFF">地址</td>
    <td width="100" bgcolor="#FFFFFF">电话</td>
    <td width="150" bgcolor="#FFFFFF">传真</td>
    <td width="80" bgcolor="#FFFFFF">法定代理人</td>
    <td width="150" bgcolor="#FFFFFF">简介</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr> 
    <td height="25" bgcolor="#FFFFFF"><a href="art_business_edit.asp?art_business_id=<%=rs("art_business_id")%>"><%=rs("allname")%></a>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("addr")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("tel")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("fax")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("legal_representative")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("memo")%>&nbsp;</td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
  <tr bgcolor="#FFFFFF"> 
    <td height="35" colspan="8">&nbsp;</td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
</body>
</html>
