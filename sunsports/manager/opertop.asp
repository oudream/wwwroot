<!--#include file="../conn/conn.asp"-->
<html>
<head>
<title>Managed</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<script language="javascript">

function editit(myform)
{
    myform.action="scheduleedit.asp";
    myform.submit();
}

function addit(myform)
{
    myform.action="scheduleadd.asp";
    myform.submit();
}


function delit(myform)
{
    myform.action="scheduledel.asp";
    myform.submit();
}

</script>

</head>
<body topmargin="3">


<table width=100% border=0 cellspacing=0 cellpadding=0>
  <tr> 
    <td  align="left"><font face="Arial"><b><font size="2">��ӭ<font color="#FF6600"><%=session("ename")%></font>��¼�ճ̰���ϵͳ</font></b></font></td>
    <td  align="right"><a href="index.asp" ><font color="#000000" size="2">������ҳ</font></a></td>
  </tr>
  <tr> 
    <td align="center" colspan=2>
<form method="get" target="mainFrame">
		 <b><font size="2">YOU WANT ...--&gt;</font></b> 
      <font size="2">&nbsp;&nbsp; 
      <input type="submit" value="DELETE" name="navdel" onClick="delit(this.form)"  >
      &nbsp;&nbsp; 
      <input type="submit" value="ADD" name="navadd" onClick="addit(this.form)" >
        &nbsp;&nbsp; </font> 
      </form>
</td>
  </tr>
</table>
	 </body>
</html>