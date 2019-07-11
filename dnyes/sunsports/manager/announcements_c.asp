<script language="JavaScript">

function checkform()
{
	if (form_administrator.inserttime.value.length == 0) {
		alert("Please enter a valid Datetime. ");
		form_administrator.inserttime.focus();
		return false;
		}
	if (form_administrator.content.value.length == 0) {
		alert("Please enter a valid content. ");
		form_administrator.content.focus();
		return false;
		}
	if(! isLetter(form_administrator.content.value)) {
		alert("Please enter a valid content. ");
		form_administrator.content.focus();
		return false;
		}
	return true;
}

function isLetter(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charCodeAt(i) > 128)
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
<!--#include file="adminconn.asp" -->
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
	
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><h2><font color="#000066"><br>
              Creat Announcements <br>
        </font></h2></td>
  </tr>
</table>

<%
if request("option")="add" then
inserttime=trim(request("inserttime"))
inserttime=replace(inserttime,"'","''")
content=trim(request("content"))
content=replace(content,"'","''")
sql="insert into announcements (pid,name,content,inserttime) values ("&session("user_pid")&",'"&session("user_name")&"','"&content&"','"&inserttime&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Your Content is send , it will be show at homepage . ');</SCRIPT>")
end if
%>
<FORM action="announcements_c.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
                        
  <TABLE width="511" border=0 cellPadding=5 
                        cellSpacing=1 bgColor=#CCCCCC>
    <TBODY>
      <TR bgColor=#ffffff>
        <TD align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>Datetime ：</TD>
        <TD bgColor=#ffffff><input name="inserttime" type="text" id="inserttime" size="30" maxlength="50" value="<%=now()%>" class="form"></TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD 
                              width="29%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>Content 
          ：</TD>
        <TD bgColor=#ffffff width="71%"><textarea cols=49 name=content id="content" rows=10></textarea> 
        </TD>
      </TR>
      <TR bgColor=#ffffff> 
        <TD bgColor=#f5f5f5 borderColor=#c9c9c9>&nbsp;</TD>
        <TD bgColor=#ffffff> <input type="hidden" id="option" name="option" value="add"> 
          <input type=submit id="options" name="options" value="OK AND SENT &gt;&gt;"> 
        </TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
</td>
  </tr>
</table>