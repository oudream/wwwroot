<script language="JavaScript">

function addnum(value_1,value_2,selObj){ 
selObj.value=value_1*value_2;
}

function checkform()
{
	if (form_administrator.gid.value.length == 0) {
		alert("��ѡȡ�ͻ�");
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

	if(! isNumber(form_administrator.mount.value)) {
		alert("��������ȷ������.");
		form_administrator.mount.focus();
		return false;
		}
	if(! isNumber(form_administrator.price.value)) {
		alert("��������ȷ�ĵ���.");
		form_administrator.price.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter.value)) {
		alert("��������ȷ�Ľ��.");
		form_administrator.counter.focus();
		return false;
		}
		
	if (form_administrator.operationer_id.value == 0) {
		alert("��ѡȡ�����Ʒ�Ĳ��������� ");
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
subspass_id=request("subspass_id")
if subspass_id="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('����ѡ����Ҫ�޸ĵ���Ʒ��ͨ! ');</SCRIPT>")
response.End()
end if
%>

<%
if request("option")="edit" then

gid=request("gid")
gcontacter=request("gcontacter")
if gcontacter="" then gcontacter=" "
pyear=request("pyear")
pmonth=request("pmonth")
pday=request("pday")
mount=request("mount")
price=request("price")
counter=request("counter")
memo=request("memo")
operationer_id=request("operationer_id")
if operationer_id="" then operationer_id=0
user_id=session("user_id")
if memo="" then memo=" "

conn.Execute("update subspass set gid="&gid&",gcontacter='"&gcontacter&"',pyear="&pyear&",pmonth="&pmonth&",pday="&pday&",mount="&mount&",price="&price&",counter="&counter&",memo='"&memo&"',operationer_id="&operationer_id&",user_id="&user_id&" where subspass_id=" & subspass_id )
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' �޸���Ʒ��ͨ�ɹ� ');</SCRIPT>")
end if
%>

<%
sql="select * from subspass where subspass_id="&subspass_id
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.eof then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('����ѡ����Ҫ�޸ĵ���Ʒ��ͨ! ');</SCRIPT>")
response.End()
end if
%>
<%
if rs("subspass_type")="in" then 
passstr="��"
passstr1="��"
end if
if rs("subspass_type")="out" then 
passstr="��"
passstr1="��"
end if
%>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="subs_pass_edit.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE cellSpacing=0 cellPadding=0 width=600 border=0>
    <TBODY>
      <TR align="center">
        <TD height="60" colSpan=2><br><font size="6" face="����">�� �� <%=passstr%> �� ��</font> 
        </TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2><table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr> 
              <td width="50%"><font color="red" size="3"><%=passstr1%></font><font size="3">����λ :
			  <select name="gid" id="gid" class="display_dropdown">
          <%
psql="select * from guest where gid="&rs("gid")
Set prs=Server.CreateObject("ADODB.RecordSet")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
%>
          <option value=<%=prs("gid")%> selected><%=prs("simname")%></option>
<%
end if
prs.close
%>
<%
psql="select * from guest where gid<>"&rs("gid")
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
%>
        </select>
</font></td>
              <td width="50%" align="right"><table width="300" border="0" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td align="right">
					  <input name="pyear" id="pyear" type="text" size="5" maxlength="4" value=<%=rs("pyear")%>>
                      �� 
					  <input name="pmonth" id="pmonth" type="text" size="3" maxlength="2" value=<%=rs("pmonth")%>>
                      �� 
                      <input name="pday" id="pday" type="text" size="3" maxlength="2" value=<%=rs("pday")%>>
                      ��
					</td>
                  </tr>
                </table></td>
            </tr>
          </table></TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2><table width="100%" border="1" cellspacing="0" cellpadding="0">
            <tr align="center"> 
              <td width="180" height="30">��Ʒ���Ƽ��ͺ�/���</td>
              <td width="50">��λ</td>
              <td width="80">����</td>
              <td width="80">����</td>
              <td width="80">���</td>
              <td width="130">��ע</td>
            </tr>
            <tr align="center"> 
              <td>
<%
Set srs=Server.CreateObject("ADODB.RecordSet")
Set brs=Server.CreateObject("ADODB.RecordSet")
ssql="select * from subs where subs_id="&rs("subs_id")
srs.Open ssql,conn,1,1
bsql="select bname from brand where brand_id="&srs("brand_id")
brs.Open bsql,conn,1,1
%>
<%=brs("bname")&" "&srs("code")&" "%>
</td>
<td><%=srs("suit")%></td>
<%
brs.close
srs.close
set brs=nothing
set srs=nothing
%>
              <td><input name="mount" id="mount" type="text" size="8" maxlength="8" value=<%=rs("mount")%>></td>
              <td><input name="price" id="price" type="text" size="8" maxlength="8" value=<%=rs("price")%>></td>
              <td><input name="counter" id="counter" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount.value,form_administrator.price.value,this);" value=<%=rs("counter")%>></td>
              <td><input name="memo" id="memo" type="text" size="15" maxlength="100" value="<%=rs("memo")%>"></td>
            </tr>
          </table>
          <font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="35%"><font color="red" size="3"><%=passstr1%></font><font size="3">����λ������
                <input name="gcontacter" id="gcontacter" type="text" size="10" maxlength="10" value="<%=rs("gcontacter")%>"></td>
              <td width="40%">����˾<%=passstr%>��������
			  <select name="operationer_id" id="operationer_id" class="display_dropdown">
          <%
psql="select * from user where id="&rs("operationer_id")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
%>
          <option value=<%=prs("id")%> selected><%=prs("uname")%></option>
<%
end if
prs.close
%>
<%
psql="select * from user where id<>"&rs("operationer_id")
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
psql="select * from user where id="&rs("user_id")
prs.Open psql,conn,1,1
zuser_name=prs("uname")
prs.close
set prs=nothing
%>
        </select>
                </td>
              <td width="25%"> 
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="20">�޸���
                      <input name="tablewrite" id="tablewrite" type="text" size="10" maxlength="10" value="<%=session("user_name")%>" disabled></td>
                  </tr>
                  <tr>
                    <td height="20">����ˣ�<%=zuser_name%> </td>
                  </tr>
                </table></td>
            </tr>
          </table></TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2> <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2> 
            <INPUT type=submit id="option" name="option" value="edit">
			<input type="hidden" id="subspass_id" name="subspass_id" value=<%=subspass_id%>>
			&nbsp;&nbsp;&nbsp;<input type="button" name="cancelit" id="cancelit" value="����ѡȡ��Ʒ��ͨ" onClick="location='subs_pass_view.asp'">
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
