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
<LINK href=site.css rel=stylesheet>
<title></title>
</head>
<body topmargin="0">
<!--#include file="dbm_adminconn.asp" -->






<%
if request("option")="添加" then

htitle=trim(request("htitle"))
htitle=replace(htitle,"'","''")
web_addr=trim(request("web_addr"))
web_addr=replace(web_addr,"'","''")
brief=trim(request("brief"))
brief=replace(brief,"'","''")
if brief="" then brief=" "
sql="insert into dbm_hildden (user_id,htitle,web_addr,brief,inserttime) values ("&session("user_id")&",'"&htitle&"','"&web_addr&"','"&brief&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('添加成功');</SCRIPT>")
end if
%>
<table width="75%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<TABLE width=100% 
                  border=1 cellPadding=3 cellSpacing=0 bordercolor="#CCCCCC">
  <FORM action="dbm_hildden_add.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
    <TBODY>
      <TR align="center"> 
        <TD height=30 colspan="2"><font color="#FF0000" size="+1"><strong> 收藏夹添加 
          </strong></font></TD>
      </TR>
      <TR> 
        <TD width="46%" align="right"> 标题：</TD>
        <TD> <input maxlength=100 size=30 id="htitle" name="htitle">
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> 网址：</TD>
        <TD> http://<INPUT maxLength=100 size=30 id="web_addr" name="web_addr"> </TD>
      </TR>
      <TR> 
        <TD align="right"> 简介：</TD>
        <TD> <textarea name="brief" id="brief" cols="60" rows="10"></textarea> 
        </TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2> <INPUT type=submit id="option" name="option" value="添加"> 
        </TD>
      </TR>
    </TBODY>
  </FORM>
</TABLE>
</body>
</html>
