<script language="JavaScript">

function checkform(zwho)
{
	if(! isNumber(zwho.value)) {
		alert("请输入正确的初始值.");
		zwho.focus();
		return false;
		}
	return true;
}

function isNumber(name)
{
	if (name.length == 0) 
		return false;
	for(i = 0; i < name.length; i++) {
		if (!((name.charAt(i) >= "0" && name.charAt(i) <= "9") || name.charAt(i) == "." || name.charAt(i) == "-"))
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

zini=request("ini")
if zini="" then zini=0
conn.Execute("update guest set ini="&zini&", user_id="&session("user_id")&" where gid=" & request("gid"))

response.Write("<SCRIPT LANGUAGE=JavaScript> alert('编辑成功');</SCRIPT>")
end if
%>
<%
sql="select * from guest order by gid" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>金钱初始化编缉</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
  <TABLE height=95 cellSpacing=0 cellPadding=0 width=550 
                  border=0>
    <TBODY>
      <TR> 
        <TD width="546" height=45> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="20%" height="30"><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">客户简称</font></td>
            <td width="20%"><font 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >初始值</font> </td>
            <td width="20%">填表人</td>
            <td width="20%">修改人</td>
            <td width="20%">&nbsp;</td>
          </tr>
        </table></TD>
      </TR>
      <TR> 
        <TD height=25> 
<%
Set prs=Server.CreateObject("ADODB.RecordSet") 
do while not rs.eof
psql="select * from user where id="&rs("user_id")
prs.Open psql,conn,1,1
if not prs.eof then 
zuser_name=prs("uname")
else
zuser_name=""
end if
prs.close
%>
<FORM action="money_guest.asp" method="post" name="form_administrator" id="form_administrator">
 <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="20%"><%=rs("simname")%></td>
              <td width="20%"><INPUT maxLength=8 size=10 id="ini" name="ini" value=<%=rs("ini")%> onBlur="return checkform(this);"></td>
              <td width="20%"><input name="tablewrite" id="tablewrite" type="text" size="10" maxlength="10" value="<%=zuser_name%>" disabled></td>
              <td width="20%"><input name="tablewrite" id="tablewrite" type="text" size="10" maxlength="10" value="<%=session("user_name")%>" disabled></td>
              <td width="20%"><INPUT type=submit id="option" name="option" value="edit">&nbsp;&nbsp;&nbsp;</td>
            </tr>
          </table>
<input type="hidden" id="gid" name="gid" value=<%=rs("gid")%>>
</FORM>
<%
rs.movenext
loop
set prs=nothing
%> </TD>
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
