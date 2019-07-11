<script language="JavaScript">
<!--
function gettarg(targ,selObj,restore){ //v3.0
  eval(targ+".location='dbm_user_guest.asp?user_id="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="dbm_adminconn.asp" -->
<%
user_id=request("user_id")
Set crs=Server.CreateObject("ADODB.RecordSet")

if trim(request("option"))="确认修改" then
guest_id_sum=request("guest_id_sum")

for j=0 to guest_id_sum-1
guest_id=request("guest_id_"&j)
guest_id_v=request("guest_id_v_"&j)
if guest_id_v=1 then
csql="select * from dbm_ug where user_id=" & user_id & " and guest_id =" & guest_id
crs.Open csql,conn,1,1
if crs.eof or crs.bof then
crs.close
conn.execute "insert into dbm_ug (user_id,guest_id) values ("&user_id&","&guest_id&")"
else
crs.close
end if
else
csql="select * from dbm_ug where user_id=" & user_id & " and guest_id =" & guest_id
crs.Open csql,conn,1,1
if not ( crs.eof or crs.bof ) then
crs.close
conn.execute "delete * from dbm_ug where user_id=" & user_id & " and guest_id =" & guest_id
else
crs.close
end if
end if

next

response.Write("<SCRIPT LANGUAGE=JavaScript> alert('修改成功');</SCRIPT>")

end if
%>
<%
if user_id="" then
csql="select * from dbm_user where adminlevel=2 order by user_id"
crs.Open csql,conn,1,1
if crs.eof or crs.bof then
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td height="10">&nbsp;</td>
  </tr>
  <tr> 
    <td height="30" align="center" valign="middle" class="header">没有一般的管理员</td>
  </tr>
</table>
<%
crs.close
response.End()
else
user_id=crs("user_id")
crs.close
end if
end if

csql="select * from dbm_user where user_id="&user_id
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
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="10">&nbsp;</td>
        </tr>
        <tr> 
          <td height="30" align="center" valign="middle" class="header">一般管理员 
            -- <font color="#FF0000"><%=allname%></font> --</td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td> 
<!-- ========================================================================================================		  

													cup table ============star

 ======================================================================================================== -->
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td> 
<form name="form1" method="post" action="dbm_user_guest.asp">
            <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
              <tr> 
                <td height="20" colspan="3" align="center" bgcolor="#FFFFFF"><table width="100%" height="100%" cellpadding=0 cellspacing=0>
                    <tbody>
                      <tr> 
                        <td width="50%" align="right"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        size=2>选择你想编缉的一般管理员</font></b></td>
                        <td width=50% align="left"> <select name="user_id" class="display_dropdown" id="user_id" onchange="gettarg('self',this,0)">
                            <option selected value=<%=user_id%>><%=allname%> </option>
                            <%
sql="select * from dbm_user where adminlevel=2 and user_id<>"&user_id&" order by user_id" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
                            <option value=<%=rs("user_id")%>><%=rs("allname")%> </option>
                            <%
rs.movenext
loop
rs.close
set rs=nothing
%>
                          </select> </td>
                      </tr>
                    </tbody>
                  </table></td>
              </tr>
<%
dim abrand()
tsql="select guest_id from dbm_ug where user_id="&user_id
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if not ( trs.eof or trs.bof ) then
redim abrand(trs.recordcount-1)
for i=0 to ubound(abrand)
abrand(i)=trs("guest_id")
trs.movenext
next
trs.close
set trs=nothing
isempty_abrand=0
else
trs.close
set trs=nothing
isempty_abrand=1
end if
%>
              <tr> 
                <td width="781" height="30" bgcolor="#FFFFFF">客户及文件夹名称</td>
                  <td width="207" bgcolor="#FFFFFF">是否可以管理</td>
              </tr>
<%
tsql="select * from dbm_guest order by allname"
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if trs.eof or trs.bof then
response.Write("<tr><td rowspan=5 align=center bgcolor=#FFFFFF>还没有任何客户资料</td></tr></table>")
response.End()
trs.close
set trs=nothing
else
j=0
do while not trs.eof
zstrtf=""
if isempty_abrand=0 then
for i=0 to ubound(abrand)
if abrand(i)=trs("guest_id") then zstrtf="checked"
next
end if
%>
              <tr> 
                <td height="25" bgcolor="#FFFFFF"><%=trs("allname")%> &nbsp;<input name="guest_id_<%=j%>" id="guest_id_<%=j%>" type="hidden" value=<%=trs("guest_id")%>></td>
                <td bgcolor="#FFFFFF">
<input name="guest_id_v_<%=j%>" id="guest_id_v_<%=j%>" type="checkbox" value=1 <%=zstrtf%>></td>
              </tr>
              <%
j=j+1
trs.movenext
loop
end if
%>
              <tr> 
                  <td height="40" colspan="3" align="center" bgcolor="#FFFFFF">
				  <input type="submit" name="option" id="option" value="确认修改">
				  <input name="guest_id_sum" id="guest_id_sum" type="hidden" value=<%=j%>>
				  </td>
              </tr>
              <%
trs.close
set trs=nothing
%>
            </table>
                  </form>
</td>
        </tr>
      </table> 
      <!-- ========================================================================================================		  

													cup table ============stop

 ======================================================================================================== -->
    </td>
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
