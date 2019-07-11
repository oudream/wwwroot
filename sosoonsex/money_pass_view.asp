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
	return true;
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
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>

<!--#include file="adminconn.asp" -->
<%
if trim(request("option"))="del" then
conn.execute "delete * from moneypass where moneypass_id=" & request("moneypass_id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('删除成功');</SCRIPT>")
end if
%>
<form name="form_administrator" id="form_administrator" method="post" action="money_pass_view.asp" onSubmit="return checkform();">
        <table width="600" border="1" align="center" cellpadding="0" cellspacing="0">
          <tr>
            
      <td width="120" height="31">流通类型</td>
            <td width="120">客户</td>
            <td width="150">日期</td>
            <td width="120">本公司经手人</td>
            <td width="90">&nbsp;</td>
          </tr>
          <tr> 
            <td width="120"> <select name="mptid" id="mptid" class="display_dropdown">
                  <option value="" selected> 请选择……</option>
                  <%
psql="select * from moneypasstype"
Set prs=Server.CreateObject("ADODB.RecordSet")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
                  <option value=<%=prs("mptid")%>><%=prs("mptname")%></option>
                  <%
prs.movenext
loop
end if
prs.close
%>
                </select></td>
            
      <td width="120"> <select name="gid" id="gid" class="display_dropdown">
          <option value=0 selected> 请选择……</option>
          <%
gsql="select * from guest"
Set grs=Server.CreateObject("ADODB.RecordSet")
grs.Open gsql,conn,1,1
if not ( grs.eof or grs.bof ) then
do while not grs.eof
%>
          <option value=<%=grs("gid")%>><%=grs("simname")%></option>
          <%
grs.movenext
loop
end if
grs.close
set grs=nothing
%>
        </select></td>
            <td width="150"> <input name="pmonth" id="pmonth" type="text" size="3" maxlength="2">
              月 
              <input name="pday" id="pday" type="text" size="3" maxlength="2">
              日 </td>
            <td width="120"><select name="operationer_id" id="operationer_id" class="display_dropdown">
                <option value=0 selected> 请选择……</option>
                <%
psql="select * from user"
Set prs=Server.CreateObject("ADODB.RecordSet")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
                <option value=<%=prs("id")%>><%=prs("uname")%></option>
                <%
prs.movenext
loop
end if
prs.close
%>
              </select></td>
            <td width="90"> <input type="submit" name="Submit" value="确定"></td>
          </tr>
        </table>
      </form> 

<%
mptid=request("mptid")
if mptid="" then mptid=0
gid=request("gid")
if gid="" then gid=0
pmonth=request("pmonth")
pday=request("pday")
operationer_id=request("operationer_id")
if operationer_id="" then operationer_id=0

sql="select * from moneypass where moneypass_id<>0" 
if mptid<>0 then sql=sql&" and mptid="&mptid
if gid<>0 then sql=sql&" and gid="&gid
if pmonth<>"" then sql=sql&" and pmonth="&pmonth
if pday<>"" then sql=sql&" and pday="&pday
if operationer_id<>0 then sql=sql&" and operationer_id="&operationer_id
sql=sql&" order by pday desc,pmonth desc,gid desc,mptid desc,operationer_id desc"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('没有所要的金钱流通');</SCRIPT>")
else
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=20
rs.AbsolutePage=pagecount
%>
<table width="600" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr bgcolor="#FFFFFF"> 
    <td height="22" colspan="9" align="center" class="header"><table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr align="center"> 
          <td height="38" colspan="6">所要查询的<font color="#FF0000">结果</font></td>
        </tr>
      </table></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="22" colspan="9" align="center" class="header"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="60%"> <table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
for zi=1 to rs.pagecount
%>
                <td width="50"><a href="money_pass_view.asp?page=<%=zj%>&mptid=<%=mptid%>&gid=<%=gid%>&pmonth=<%=pmonth%>&pday=<%=pday%>&operationer_id=<%=operationer_id%>">&nbsp;<%=zj%>&nbsp;</a></td>
                <%
zj=zj+1
next
%>
              </tr>
            </table></td>
          <td width="24%" align="left"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td> <div align="center"> Record :<font color=red><b><%=rs.recordcount%></b></font> Page <font color=red><b><%=pagecount%></b></font> of <font color=red><b><%=rs.pagecount%></b></font> </div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="90" height="30" bgcolor="#FFFFFF">流通类型</td>
    <td width="50" bgcolor="#FFFFFF">金额</td>
    <td width="100" bgcolor="#FFFFFF">日期</td>
    <td width="80" bgcolor="#FFFFFF">经手人</td>
    <td width="80" bgcolor="#FFFFFF">客户</td>
    <td width="120" bgcolor="#FFFFFF">备注</td>
    <td width="80" bgcolor="#FFFFFF">Operation</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr> 
    <td width="90" height="25" bgcolor="#FFFFFF"><%
psql="select * from moneypasstype where mptid="&rs("mptid")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
%>
<%=prs("mptname")%>
<%
end if
prs.close
%>
&nbsp;</td>
    <td width="50" bgcolor="#FFFFFF"><%=rs("counter")%>&nbsp;</td>
    <td width="100" bgcolor="#FFFFFF"><%=rs("pmonth")%> 月 <%=rs("pday")%>日 </td>
    <td width="80" bgcolor="#FFFFFF"> <%
if rs("operationer_id")<>0 then
usql="select * from user where id=" & rs("operationer_id")
Set urs=Server.CreateObject("ADODB.RecordSet") 
urs.open usql,conn,1,1 
if not ( urs.bof or urs.eof ) then 
response.Write(urs("uname"))
end if
urs.close
end if
%> 
      &nbsp;</td>
    <td width="80" bgcolor="#FFFFFF">
<%
Set zrs=Server.CreateObject("ADODB.RecordSet")
zsql="select simname from guest where gid="&rs("gid")
zrs.Open zsql,conn,1,1
if not ( zrs.bof or zrs.eof ) then
response.Write(zrs("simname"))
end if
zrs.close
set zrs=nothing
%>
	&nbsp;</td>
    <td width="120" bgcolor="#FFFFFF"><%=rs("memo")%>&nbsp;</td>
    <td width="80" bgcolor="#FFFFFF"><a href="money_pass_edit.asp?moneypass_id=<%=rs("moneypass_id")%>">修改</a>&nbsp;&nbsp;&nbsp;<a href="money_pass_view.asp?option=del&moneypass_id=<%=rs("moneypass_id")%>" onClick="return confirm('确定删除此项吗？')">删除</a></td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
  <tr bgcolor="#FFFFFF"> 
    <td height="35" colspan="9">&nbsp;</td>
  </tr>
</table>
<%
end if
rs.close
set rs=nothing
%>
</body>
</html>
