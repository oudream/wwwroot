<script language="JavaScript">

function checkform()
{
	if (form_administrator.sname.value.length == 0) {
		alert("请输入种类名称");
		form_administrator.sname.focus();
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

sname=trim(request("sname"))
sname=replace(sname,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")
if memo="" then memo=" "
usql="select * from sort where sname='"&sname&"' and sort_id<>" & request("sort_id")
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('出错！此种类有同名存在，请用其它名字输入');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing
conn.Execute("update sort set sname='"&sname&"', memo='"&memo&"' where sort_id=" & request("sort_id"))
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('编辑成功');</SCRIPT>")
end if
end if
%>
<%
sql="select * from sort where sort_id="&request("sort_id") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('没有所需要的种类');</SCRIPT>")
response.End()
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>种类编缉</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="sort_edit.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE height=127 cellSpacing=0 cellPadding=0 width=550 
                  border=0>
    <TBODY>
      <TR> 
        <TD width=150 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">种类名称</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="sname" name="sname" value="<%=rs("sname")%>">
          * </FONT></TD>
      </TR>
      <TR> 
        <TD width="150" height=32> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">备注</font></DIV></TD>
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
            <input type="hidden" id="sort_id" name="sort_id" value="<%=request("sort_id")%>">
            </FONT></STRONG></P></TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
<br>
<br>
</body>
</html>
