<script language="JavaScript">

function checkform()
{

	if (form_administrator.pmonth.value.length > 0) {
	if(! TFNumber(form_administrator.pmonth.value)) {
		alert("请用数字填写月份");
		form_administrator.pmonth.focus();
		return false;
		}
	if (form_administrator.pmonth.value > 12 || form_administrator.pmonth.value < 1) {
		alert("请输入正确的月份, 必须在 1 至 12 . ");
		form_administrator.pmonth.focus();
		return false;
		}
		}

	if (form_administrator.pday.value.length > 0) {
	if(! TFNumber(form_administrator.pday.value)) {
		alert("请用数字填写日期");
		form_administrator.pday.focus();
		return false;
		}
	if (form_administrator.pday.value > 31 || form_administrator.pday.value < 1) {
		alert("请输入正确的日期, 必须在 1 至 31 . ");
		form_administrator.pday.focus();
		return false;
		}
		}
}		

function TFNumber(name)
{
	for(i = 0; i < name.length; i++) {
		if (name.charAt(i) < "0" || name.charAt(i) > "9")
			return false;
	}
	return true;
}

</script>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="8" topmargin="8" marginwidth="0" marginheight="0">
<table width="600" height="30" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="30">&nbsp;</td>
  </tr>
</table>
<!--#include file="adminconn.asp" -->
<%
pmonth=request("pmonth")
if pmonth="" then pmonth=12
pday=request("pday")
if pday="" then pday=31

dim aini()
dim asum_in()
dim asum_out()
dim asum_money()
dim asum_all()
sql="select * from guest order by simname" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
%>
<table width="600" height="50" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="200" height="50"><font size="5"><strong>数信债权状况</strong></font></td>
    <td width="400" height="50" align="right">
<FORM action="money_base_view.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
      日期 
      <input name="pmonth" id="pmonth" type="text" size="3" maxlength="2">
        月 
        <input name="pday" id="pday" type="text" size="3" maxlength="2">
        日 
		<input type="submit" name="Submit" value="查询>>">
      </form>
		</td>
  </tr>
</table>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
  <TABLE height=95 cellSpacing=0 cellPadding=0 width=600 
                  border=0>
    <TBODY>
      <TR> 
        <TD width="546" height=45> <table width="600" border="1" cellspacing="0" cellpadding="0">
      <TR align="center"><TD height=25 colspan="6"><font color="#FF0000">
<%
if pmonth=12 and pday=31 then
response.write("目前")
else
response.write(pmonth&" 月 "&pday&" 日 ")
end if
%>
              </font>的债权状况 （是计算此日期、此日期之前的）</TD>
          </TR>
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

qsql="SELECT sum(counter) as sum_in from subspass where subspass_type='in' and gid="&rs("gid")&" and pmonth <=" & pmonth&" and pday <=" & pday
qrs.Open qsql,conn,1,1
sum_in=0-qrs("sum_in")
if isnull(qrs("sum_in")) then sum_in=0
asum_in(i)=sum_in
qrs.close

qsql="SELECT sum(counter) as sum_out from subspass where subspass_type='out' and gid="&rs("gid")&" and pmonth <=" & pmonth&" and pday <=" & pday
qrs.Open qsql,conn,1,1
sum_out=qrs("sum_out")
if isnull(qrs("sum_out")) then sum_out=0
asum_out(i)=sum_out
qrs.close

qsql="SELECT sum(counter) as sum_money from moneypass where gid="&rs("gid")&" and pmonth <=" & pmonth&" and pday <=" & pday
qrs.Open qsql,conn,1,1
sum_money=qrs("sum_money")
if isnull(qrs("sum_money")) then sum_money=0
asum_money(i)=sum_money
qrs.close

asum_all(i)=rs("ini")+sum_in+sum_out-sum_money
%>
          <tr> 
            <td width="20%" height="26"><a href="guest_edit.asp?gid=<%=rs("gid")%>"><%=rs("simname")%><a></td>
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
