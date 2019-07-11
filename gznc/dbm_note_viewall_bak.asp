<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--media=print 这个属性可以在打印时有效--> 
<style media=print> 
.Noprint{display:none;} 
.PageNext{page-break-after: always;} 
</style> 

<style> 
.tdp 
{ 
border-bottom: 1 solid #000000; 
border-left: 1 solid #000000; 
border-right: 0 solid #ffffff; 
border-top: 0 solid #ffffff; 
} 
.tabp 
{ 
border-color: #000000 #000000 #000000 #000000; 
border-style: solid; 
border-top-width: 2px; 
border-right-width: 2px; 
border-bottom-width: 1px; 
border-left-width: 1px; 
} 
.NOPRINT { 
font-family: "宋体"; 
font-size: 9pt; 
} 

</style>
<title></title></head>
<body topmargin="0">
<!--#include file="dbm_adminconn.asp" -->


  <table width="550" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>
	  <center class="Noprint" >
    <OBJECT id=WebBrowser classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 height=0 width=0>
    </OBJECT>
    <input type=button value=打印 onclick=document.all.WebBrowser.ExecWB(6,1)>&nbsp;
    <input type=button value=直接打印 onclick=document.all.WebBrowser.ExecWB(6,6)>&nbsp;
    <input type=button value=页面设置 onclick=document.all.WebBrowser.ExecWB(8,1)>&nbsp;
    <input name="button" type=button onClick=document.all.WebBrowser.ExecWB(7,1) value=打印预览>
      </center> 
</td>
    </tr>
  </table>

<%
sql="select * from dbm_note where id="&request("zid") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('没有所需要的');</SCRIPT>")
response.End()
end if
%>
<table width="650" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
  <TABLE width=650 
                  border=1 align="center" cellPadding=3 cellSpacing=0 bordercolor="#CCCCCC">
    <TBODY>
      <TR align="center"> 
        <TD height="30" colspan="2"><strong><font color="#FF0000"> 留言编缉</DIV></font></strong></TD>
      </TR>
      <TR>
        <TD height=11 colspan="2">&nbsp;</TD>
      </TR>
      <TR> 
        <TD width="46%" align="right">留言人：</TD>
        <TD> <%
usql="select * from dbm_user where user_id="&rs("user_id") 
Set urs=Server.CreateObject("ADODB.RecordSet") 
urs.open usql,conn,1,1 
allname=urs("allname")
urs.close
set urs=nothing
%> 
      <%=allname%>&nbsp; </TD>
      </TR>
      <TR> 
        <TD width="46%" align="right">留言时间：</TD>
        <TD> 
         <%=rs("inserttime")%></TD>
      </TR>
      <TR> 
        <TD width="46%" align="right">留言主题：</TD>
        <TD> 
         <%=rs("subject")%></TD>
      </TR>
      <TR> 
        <TD align="right"> 留言：</DIV></TD>
        <TD> 
          <%=rs("memo")%></TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2>&nbsp;
            </TD>
      </TR>
    </TBODY>
  </TABLE>
<br>
<br>
</body>
</html>
