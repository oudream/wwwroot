<script language="JavaScript">

function checkform()
{
	if (form_administrator.mptname.value.length == 0) {
		alert("��������������");
		form_administrator.mptname.focus();
		return false;
		}
	if (form_administrator.sign.value.length == 0) {
		alert("��ѡȡ��������");
		form_administrator.sign.focus();
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
if request("option")="edit" then

mptname=trim(request("mptname"))
mptname=replace(mptname,"'","''")
sign=request("sign")
sign=replace(sign,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")
if memo="" then memo=" "
usql="select * from moneypasstype where mptname='"&mptname&"' and mptid<>" & request("mptid")
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('������������ͬ�����ڣ�����������������');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing
conn.Execute("update moneypasstype set mptname='"&mptname&"', sign='"&sign&"', memo='"&memo&"' where mptid=" & request("mptid"))
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('�༭�ɹ�');</SCRIPT>")
end if
end if
%>
<%
sql="select * from moneypasstype where mptid="&request("mptid") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('û������Ҫ������');</SCRIPT>")
response.End()
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>��Ǯ��ͨ���ͱ༩</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="moneypasstype_edit.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE height=127 cellSpacing=0 cellPadding=0 width=550 
                  border=0>
    <TBODY>
      <TR> 
        <TD width=150 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">��������</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="mptname" name="mptname" value="<%=rs("mptname")%>">
          * </FONT></TD>
      </TR>
      <TR> 
        <TD width="150">��������</TD>
        <TD width="71%" bgColor=#ffffff>
		<select name="sign" id="sign">
<%
if rs("sign")="plus" then
%>
            <option value="plus" selected>����</option>
            <option value="minus">֧��</option>
<%
else
%>
            <option value="plus">����</option>
            <option value="minus" selected>֧��</option>
<%
end if
%>
          </select>
          *</TD>
      </TR>
      <TR> 
        <TD width="150" height=32> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">��ע</font></DIV></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <textarea name="memo" id="memo" cols="60" rows="10"><%=rs("memo")%></textarea>
          </FONT></TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2> <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2> 
            <INPUT type=submit id="option" name="option" value="edit">
            &nbsp;&nbsp;&nbsp; 
            <INPUT type="button" id="goback" name="goback" value="go back" onClick="history.go( -1 );">
            <input type="hidden" id="mptid" name="mptid" value="<%=request("mptid")%>">
            </FONT></STRONG></P></TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
<br>
<br>
</body>
</html>
