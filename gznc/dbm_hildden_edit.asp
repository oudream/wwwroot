<script language="JavaScript">

function checkform()
{
	if (form_administrator.htitle.value.length == 0) {
		alert("请输入标题");
		form_administrator.htitle.focus();
		return false;
		}
	if (form_administrator.web_addr.value.length == 0) {
		alert("请输入网址");
		form_administrator.web_addr.focus();
		return false;
		}
	return true;
}

</script>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="dbm_css.css" type=text/css rel=stylesheet>
<title></title>
</head>
<body topmargin="0">
<!--#include file="dbm_adminconn.asp" -->






<%
if request("option")="编辑" then

htitle=trim(request("htitle"))
htitle=replace(htitle,"'","''")
web_addr=trim(request("web_addr"))
web_addr=replace(web_addr,"'","''")
brief=trim(request("brief"))
brief=replace(brief,"'","''")
if brief="" then brief=" "
conn.Execute("update dbm_hildden set htitle='"&htitle&"',web_addr='"&web_addr&"',brief='"&brief&"' where dbm_hildden_id=" & request("dbm_hildden_id"))
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('编辑成功');</SCRIPT>")
end if
%>
<%
sql="select * from dbm_hildden where dbm_hildden_id="&request("dbm_hildden_id") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' None data ! ');</SCRIPT>")
response.End()
end if
%>
<table width="75%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<TABLE width=100% 
                  border=1 cellPadding=3 cellSpacing=0 bordercolor="#CCCCCC">
  <FORM action="dbm_hildden_edit.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
    <TBODY>
      <TR align="center"> 
        <TD height=30 colspan="2"><font color="#FF0000" size="+1"><strong>网址编缉</strong></font></TD>
      </TR>
      <TR> 
        <TD width="46%" align="right"> 标题：</TD>
        <TD> <input maxlength=100 size=30 id="htitle" name="htitle" value="<%=rs("htitle")%>">
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> 网址：</TD>
        <TD> <INPUT maxLength=40 size=30 id="web_addr" name="web_addr" value="<%=rs("web_addr")%>"> 
        </TD>
      </TR>
      <TR> 
        <TD align="right"> 简介：</TD>
        <TD> <textarea name="brief" id="brief" cols="60" rows="10"><%=rs("brief")%></textarea> 
        </TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2> <INPUT type=submit id="option" name="option" value="编辑"> 
          &nbsp;&nbsp;&nbsp; <INPUT type="button" id="goback" name="goback" value="go back" onClick="history.go( -1 );"> 
          <input type="hidden" id="dbm_hildden_id" name="dbm_hildden_id" value="<%=request("dbm_hildden_id")%>"> 
        </TD>
      </TR>
    </TBODY>
  </FORM>
</TABLE>
</body>
</html>
