<script language="JavaScript">

function addnum(value_1,value_2,selObj){ 
selObj.value=value_1*value_2;
}

function checkform()
{
	if (form_administrator.mptid.value.length == 0) {
		alert("��ѡȡ��ͨ����");
		form_administrator.mptid.focus();
		return false;
		}
	if (form_administrator.gid.value.length == 0) {
		alert("��ѡȡ�����Ʒ������ ");
		form_administrator.gid.focus();
		return false;
		}
	if (form_administrator.pyear.value.length == 0) {
		alert("��������� ");
		form_administrator.pyear.focus();
		return false;
		}
	if(! TFNumber(form_administrator.pyear.value)) {
		alert("����������д���");
		form_administrator.pyear.focus();
		return false;
		}
	if (form_administrator.pyear.value > 2005 || form_administrator.pyear.value < 2003) {
		alert("��������ȷ�����, ������ 2004 �� 2005 . ");
		form_administrator.pyear.focus();
		return false;
		}
		
	if (form_administrator.pmonth.value.length == 0) {
		alert("�������·� ");
		form_administrator.pmonth.focus();
		return false;
		}
	if(! TFNumber(form_administrator.pmonth.value)) {
		alert("����������д�·�");
		form_administrator.pmonth.focus();
		return false;
		}
	if (form_administrator.pmonth.value > 12 || form_administrator.pmonth.value < 1) {
		alert("��������ȷ���·�, ������ 1 �� 12 . ");
		form_administrator.pmonth.focus();
		return false;
		}

	if (form_administrator.pday.value.length == 0) {
		alert("���������� ");
		form_administrator.pday.focus();
		return false;
		}
	if(! TFNumber(form_administrator.pday.value)) {
		alert("����������д����");
		form_administrator.pday.focus();
		return false;
		}
	if (form_administrator.pday.value > 31 || form_administrator.pday.value < 1) {
		alert("��������ȷ������, ������ 1 �� 31 . ");
		form_administrator.pday.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter.focus();
		return false;
		}
	if (form_administrator.operationer_id.value == 0) {
		alert("��ѡ�񱾹�˾������ ");
		form_administrator.operationer_id.focus();
		return false;
		}
	return true;
}

function isNumber(name)
{
	if (name.length == 0) 
		return false;
	for(i = 0; i < name.length; i++) {
		if (!((name.charAt(i) >= "0" && name.charAt(i) <= "9") || name.charAt(i) == "."))
			return false;
	}
	return true;
}

function TFNumber(name)
{
	if (name.length == 0) 
		return false;
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
<!--#include file="adminconn.asp" -->
<%
if request("option")="add" then

mptid=request("mptid")
gid=request("gid")
pyear=request("pyear")
pmonth=request("pmonth")
pday=request("pday")
counter=request("counter")

sql="select * from moneypasstype where mptid="&mptid
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
zsign=rs("sign")
rs.close
set rs=nothing

if zsign="minus" or zsign="put" then counter=cint("-"&cstr(counter))

memo=request("memo")
if memo="" then memo=" "
gcontacter=request("gcontacter")
if gcontacter="" then gcontacter=" "
operationer_id=request("operationer_id")
if operationer_id="" then operationer_id=0
user_id=session("user_id")

sql="insert into moneypass (mptid,gid,gcontacter,pyear,pmonth,pday,counter,memo,operationer_id,user_id,inserttime) values ("&mptid&","&gid&",'"&gcontacter&"',"&pyear&","&pmonth&","&pday&","&counter&",'"&memo&"',"&operationer_id&","&session("user_id")&",'"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('��ӳɹ�');</SCRIPT>")

end if
%>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="money_pass.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE cellSpacing=0 cellPadding=0 width=600 border=0>
    <TBODY>
      <TR align="center">
        <TD height="60" colSpan=2><br><font size="6" face="����">�� �� �� Ǯ</font> 
        </TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2><table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr> 
              <td width="40%">��ͨ����
<select name="mptid" id="mptid" class="display_dropdown">
                  <option value="" selected> ��ѡ�񡭡�</option>
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
              <td width="29%">�ͻ�
<select name="gid" id="gid" class="display_dropdown">
                  <option value="" selected> ��ѡ�񡭡�</option>
                  <%
psql="select * from guest"
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
                  <option value=<%=prs("gid")%>><%=prs("simname")%></option>
                  <%
prs.movenext
loop
end if
prs.close
set prs=nothing
%>
                </select></td>
              <td width="31%" align="right"><table border="0" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td align="right"> <input name="pyear" id="pyear" type="text" size="5" maxlength="4" value="<%=year(now)%>">
                      �� 
                      <input name="pmonth" id="pmonth" type="text" size="3" maxlength="2" value="<%=month(now)%>">
                      �� 
                      <input name="pday" id="pday" type="text" size="3" maxlength="2" value="<%=day(now)%>">
                      �� </td>
                  </tr>
                </table></td>
            </tr>
          </table></TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2><table width="100%" border="1" cellspacing="0" cellpadding="0">
            <tr align="center"> 
              <td width="150" height="30">���</td>
              <td width="400">��ע</td>
            </tr>
            <tr align="center"> 
              <td height="30"><input name="counter" id="counter" type="text" size="10" maxlength="8" onFocus="addnum(form_administrator.mount_1.value,form_administrator.price_1.value,this);"></td>
              <td><input name="memo" id="memo" type="text" size="50" maxlength="100"></td>
            </tr>
          </table>
          <font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="35%" height="30">�ͻ������� 
                <input name="gcontacter" id="gcontacter" type="text" size="10" maxlength="10"></td>
              <td width="40%">����˾������<select name="operationer_id" id="operationer_id" class="display_dropdown">
                  <option value=0 selected> ��ѡ�񡭡�</option>
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
set prs=nothing
%>
                </select></td>
              <td width="25%">�����<input name="tablewrite" id="tablewrite" type="text" size="10" maxlength="10" value="<%=session("user_name")%>" disabled></td>
            </tr>
          </table></TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2> <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2> 
            <INPUT type=submit id="option" name="option" value="add">
			&nbsp;&nbsp;&nbsp;<input type="reset" name="cancelit" id="cancelit" value="���">
			&nbsp;&nbsp;&nbsp;<input type="button" name="cancelit" id="cancelit" value="ȡ�����в���" onClick="location='adminlogin.asp'">
            </FONT></STRONG></P>
          </TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
<br>
<br>
</body>
</html>
