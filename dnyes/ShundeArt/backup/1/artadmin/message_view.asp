<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>

<!--#include file="adminconn.asp" -->
<%
sql="select * from art_message order by id desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' None data ! ');</SCRIPT>")
response.End()
end if

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=10
rs.AbsolutePage=pagecount
%>
<table width="97%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="22" colspan="8" class="header"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="75%"> <table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
for zi=1 to rs.pagecount
%>
                <td width="50"><a href="message_view.asp?page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
                <%
zj=zj+1
next
%>
              </tr>
            </table></td>
          <td width="26%" align="left"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td> <div align="center"> Record :<font color=red><b><%=rs.recordcount%></b></font> 
                    Page <font color=red><b><%=pagecount%></b></font> of <font color=red><b><%=rs.pagecount%></b></font> 
                  </div></td>
              </tr>
            </table></td>
        </tr>
      </table> </td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="22" colspan="8" class="header">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="63%"><font size="5"><strong>留言列表 </strong></font></td>
          <td width="37%" align="right"><font size="4"><strong><a href="message_add.asp">留言添加</a></strong></font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="130" align="center" bgcolor="#FFFFFF">主题</td>
    <td width="70" align="center" bgcolor="#FFFFFF">留言人</td>
    <td width="80" align="center" bgcolor="#FFFFFF">留言时间</td>
    <td width="320" align="center" bgcolor="#FFFFFF">留言</td>
  </tr>
  <%
Set urs=Server.CreateObject("ADODB.RecordSet") 
do while not rs.eof 
%>
  <tr> 
    <td width="130" align="center" bgcolor="#FFFFFF"><a href="message_edit.asp?zid=<%=rs("id")%>"><%=rs("subject")%></a>&nbsp;</td>
    <td width="70" align="center" bgcolor="#FFFFFF">
<%
usql="select * from art_user where id="&rs("user_id") 
Set urs=Server.CreateObject("ADODB.RecordSet") 
urs.open usql,conn,1,1 
uname=urs("uname")
urs.close
%>
<%=uname%>&nbsp;
	</td>
    <td width="80" align="center" bgcolor="#FFFFFF"><%=rs("inserttime")%>&nbsp;</td>
    <td width="320" align="center" bgcolor="#FFFFFF"><%=rs("memo")%>&nbsp;</td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
set urs=nothing
%>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="35" colspan="8"><font size="4"><strong><a href="message_add.asp">留言添加</a></strong></font></td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
</body>
</html>
