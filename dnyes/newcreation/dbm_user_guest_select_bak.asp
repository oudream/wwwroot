<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>

<body>
<!--#include file="dbm_adminconn.asp" -->
<%
if trim(request("option"))="add" then
conn.execute "insert into dbm_ug (user_id,guest_id) values ("&request("user_id")&","&request("guest_id")&")"
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('增添成功');</SCRIPT>")
end if
%>
<%
user_id=request("user_id")
if user_id="" then user_id=1
csql="select * from dbm_user where user_id="&user_id
Set crs=Server.CreateObject("ADODB.RecordSet")
crs.Open csql,conn,1,1
if crs.eof or crs.bof then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('没有所需要的管理员');</SCRIPT>")
response.End()
crs.close
set crs=nothing
end if
allname=crs("allname")
crs.close
set crs=nothing
%>
<%
dim abrand()
tsql="select guest_id from dbm_ug where user_id="&user_id
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if not ( trs.eof or trs.bof ) then
abrandft="yes"
redim abrand(trs.recordcount-1)
for i=0 to ubound(abrand)
abrand(i)=trs("guest_id")
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
          <td height="30" align="center" valign="middle"><strong>在 -- <font color="#FF0000"><%=allname%></font> -- 增添品牌</strong></td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
        <tr> 
          <td width="80%" height="50">品牌名称</td>
          <td width="20%">选项</td>
        </tr>
        <%
tsql="select * from dbm_guest where guest_id <> 0"
if abrandft="yes" then
for i=0 to ubound(abrand)
tsql=tsql&" and guest_id <>"&abrand(i)
next
end if
tsql =tsql &" order by allname "
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if not ( trs.eof or trs.bof ) then 
do while not trs.eof
%>
        <tr> 
          <td width="80%" height="25"><%=trs("allname")%>&nbsp;</td>
          <td width="20%"><a href="dbm_user_guest_select_bak.asp?user_id=<%=user_id%>&guest_id=<%=trs("guest_id")%>&option=add">allname 可管理此客户</a></td>
        </tr>
        <%
trs.movenext
loop
end if
trs.close
set trs=nothing
%>
        <tr> 
          <td height="38" colspan="2">&nbsp;</td>
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
