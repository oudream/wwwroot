<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title></title>
</head>
<body topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>

<!--#include file="art_adminconn.asp" -->






<%
if request("option_id_del")="succ" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('删除成功');</SCRIPT>")
end if

sql="select * from art_note_old order by id desc" 
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
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="30" colspan="8" class="header"><strong><font color="#FF0000">纪事列表</font></strong></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="22" colspan="8" class="header">&nbsp; </td>
  </tr>
  <tr> 
    <td width="150" align="center" bgcolor="#FFFFFF">纪事时间</td>
    <td align="center" bgcolor="#FFFFFF">纪事</td>
  </tr>
  <%
Set urs=Server.CreateObject("ADODB.RecordSet") 
do while not rs.eof 
%>
  <tr> 
    <td align="center" bgcolor="#FFFFFF"><a href="art_note_old_edit.asp?zid=<%=rs("id")%>"><%=rs("note_time")%></a>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"><%=rs("memo")%>&nbsp;</td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
set urs=nothing
%>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="35" colspan="8"><font size="4"><strong><a href="art_note_old_add.asp">纪事添加</a></strong></font></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="75%" height="30"> <table border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <%
zj=1
for zi=1 to rs.pagecount
%>
          <td><a href="art_note_old_view.asp?page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
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
</table>
<%
rs.close
set rs=nothing
%>
</body>
</html>
