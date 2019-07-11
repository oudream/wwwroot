<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="8" topmargin="8" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" -->
<%
dim aini()
dim asum_in()
dim asum_out()
dim asum_money()
dim asum_all()
sql="select * from guest order by gid" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>数信债权状况</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
  <TABLE height=95 cellSpacing=0 cellPadding=0 width=550 
                  border=0>
    <TBODY>
      <TR> 
        <TD width="546" height=45> <table width="100%" border="1" cellspacing="0" cellpadding="0">
      <TR><TD height=25 colspan="6">&nbsp;</TD></TR>
          <tr> 
            <td width="20%" height="30"><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">客户简称</font></td>
            <td width="16%">初始值</td>
            <td width="16%">拿货</td>
            <td width="16%">出货</td>
            <td width="16%">金钱流通</td>
            <td width="16%"><strong>总</strong></td>
          </tr>
<%
redim aini(rs.recordcount-1)
redim asum_in(rs.recordcount-1)
redim asum_out(rs.recordcount-1)
redim asum_money(rs.recordcount-1)
redim asum_all(rs.recordcount-1)
i=0

Set qrs=Server.CreateObject("ADODB.RecordSet") 
do while not rs.eof

aini(i)=rs("ini")

qsql="SELECT sum(counter) as sum_in from subspass where subspass_type='in' and gid="&rs("gid")
qrs.Open qsql,conn,1,1
sum_in=0-qrs("sum_in")
if isnull(qrs("sum_in")) then sum_in=0
asum_in(i)=sum_in
qrs.close

qsql="SELECT sum(counter) as sum_out from subspass where subspass_type='out' and gid="&rs("gid")
qrs.Open qsql,conn,1,1
sum_out=qrs("sum_out")
if isnull(qrs("sum_out")) then sum_out=0
asum_out(i)=sum_out
qrs.close

qsql="SELECT sum(counter) as sum_money from moneypass where gid="&rs("gid")
qrs.Open qsql,conn,1,1
sum_money=qrs("sum_money")
if isnull(qrs("sum_money")) then sum_money=0
asum_money(i)=sum_money
qrs.close

asum_all(i)=rs("ini")+sum_in+sum_out-sum_money
%>
          <tr> 
            <td width="20%" height="25"><%=rs("simname")%></td>
            <td width="16%"><%=rs("ini")%></td>
            <td width="16%"><%=sum_in%></td>
            <td width="16%"><%=sum_out%></td>
            <td width="16%"><%=sum_money%></td>
            <td width="16%"><font color="#FF0000"><%=asum_all(i)%></font></td>
          </tr>
<%
rs.movenext
i=i+1
loop
set qrs=nothing

sum_aini=0
for i=0 to ubound(aini)
sum_aini=sum_aini+aini(i)
next

sum_asum_in=0
for i=0 to ubound(asum_in)
sum_asum_in=sum_asum_in+asum_in(i)
next

sum_asum_out=0
for i=0 to ubound(asum_out)
sum_asum_out=sum_asum_out+asum_out(i)
next

sum_asum_money=0
for i=0 to ubound(asum_money)
sum_asum_money=sum_asum_money+asum_money(i)
next

sum_asum_all=0
for i=0 to ubound(asum_all)
sum_asum_all=sum_asum_all+asum_all(i)
next
%>
          <tr> 
            <td width="20%" height="35"><strong>合计</strong></td>
            <td width="16%"><font color="#FF0000"><%=sum_aini%></font> </td>
            <td width="16%"><font color="#FF0000"><%=sum_asum_in%></font></td>
            <td width="16%"><font color="#FF0000"><%=sum_asum_out%></font></td>
            <td width="16%"><font color="#FF0000"><%=sum_asum_money%></font></td>
            <td width="16%"><font color="#FF0000"><%=sum_asum_all%></font></td>
          </tr>
      <TR><TD height=25 colspan="6">&nbsp;</TD></TR>
        </table></TD>
      </TR>
      <TR> 
        
      <TD height=35 align="center">“<font color="#FF0000"><strong>＋</strong></font>”表示数信收入；“<font color="#FF0000"><strong>－</strong></font>”表示数信支出。</TD>
      </TR>
      <TR> 
        <TD height="50"><INPUT type="button" id="goback" name="goback" value="go back" onClick="history.go( -1 );"></TD>
      </TR>
    </TBODY>
  </TABLE>
<br>
<br>
</body>
</html>
