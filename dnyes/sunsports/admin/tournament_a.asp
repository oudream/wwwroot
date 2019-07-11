<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>

<body>
<!--#include file="adminconn.asp" -->
<%
if trim(request("option"))="add" then
conn.execute "insert into ct (cid,tid,inserttime) values ("&request("cid")&","&request("tid")&",'"&now()&"')"
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' ADD TEAM TO THE TOURNAMENT is complete. ');</SCRIPT>")
end if
%>
<%
cid=request("cid")
if cid="" then cid=210
csql="select * from tournament where cid="&cid
Set crs=Server.CreateObject("ADODB.RecordSet")
crs.Open csql,conn,1,1
if crs.eof or crs.bof then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' No league you want in the databse ! ');</SCRIPT>")
response.End()
crs.close
set crs=nothing
end if
cname=crs("name")
crs.close
set crs=nothing
%>
<%
dim ateam()
tsql="select tid from ct where cid="&cid
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if not ( trs.eof or trs.bof ) then
ateamft="yes"
redim ateam(trs.recordcount-1)
for i=0 to ubound(ateam)
ateam(i)=trs("tid")
trs.movenext
next
trs.close
set trs=nothing
end if
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="10">&nbsp;</td>
        </tr>
        <tr> 
          <td height="30" align="center" valign="middle"><strong>ADD TEAM TO THE TOURNAMENT</strong></td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
        <tr>
          <td width="24%">TEAM ID</td>
          <td width="42%">TEAM NAME</td>
          <td width="34%">OPTION</td>
        </tr>
<%
tsql="select * from team where tid <> 0"
if ateamft="yes" then
for i=0 to ubound(ateam)
tsql=tsql&" and tid <>"&ateam(i)
next
tsql =tsql &" order by name "
end if
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if not ( trs.eof or trs.bof ) then 
do while not trs.eof
%>
        <tr>
          <td><%=trs("tid")%>&nbsp;</td>
          <td><%=trs("name")%>&nbsp;</td>
          <td><a href="tournament_a.asp?cid=<%=cid%>&tid=<%=trs("tid")%>&option=add">ADD TO THE TOURNAMENT</a></td>
        </tr>
<%
trs.movenext
loop
end if
trs.close
set trs=nothing
%>
        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
