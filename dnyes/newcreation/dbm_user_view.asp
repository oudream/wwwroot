<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="dbm_css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>

<!--#include file="dbm_adminconn.asp" -->


<%
if trim(request("option"))="del" then
conn.execute "delete * from dbm_user where user_id=" & request("user_id")
conn.execute "delete * from dbm_ug where user_id=" & request("user_id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('删除成功');</SCRIPT>")
end if
%>


<%
sql="select * from dbm_user order by user_id desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' None data ! ');</SCRIPT>")
response.End()
end if
%>
<table width="97%" border="1" align="center"  cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr bgcolor="#FFFFFF"> 
    <td height="22" colspan="5"  class="header"> </td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="22" colspan="5"  class="header">管理员列表</td>
  </tr>
  <tr> 
    <td width="200" height="30" bgcolor="#FFFFFF">姓名 -- 地址 -- 电话 </td>
    <td width="150"  bgcolor="#FFFFFF">级别 -- E-mail -- QQ -- MSN</td>
    <td width="150"  bgcolor="#FFFFFF">备注</td>
    <td width="100"  bgcolor="#FFFFFF"> 操作</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr> 
    <td width="200" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp; <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><%=rs("allname")%></td>
        </tr>
        <tr> 
          <td><%=rs("addr")%></td>
        </tr>
        <tr> 
          <td><%=rs("tel2")%></td>
        </tr>
        <tr> 
          <td><%=rs("tel3")%></td>
        </tr>
      </table></td>
    <td width="150"  bgcolor="#FFFFFF">&nbsp; <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td>
<%
if rs("adminlevel")=1 then response.Write("高级管理员")
if rs("adminlevel")=2 then response.Write("一般管理员")
%>
</td>
        </tr>
        <tr> 
          <td><%=rs("email")%></td>
        </tr>
        <tr> 
          <td><%=rs("qq")%></td>
        </tr>
        <tr> 
          <td><%=rs("msn")%></td>
        </tr>
      </table></td>
    <td width="150"  bgcolor="#FFFFFF"><%=rs("memo")%>&nbsp;</td>
    <td width="100"  bgcolor="#FFFFFF"><a href="dbm_user_edit.asp?user_id=<%=rs("user_id")%>">编辑</a> 
      | <a href="dbm_user_view.asp?user_id=<%=rs("user_id")%>&allname=<%=rs("allname")%>&option=del" onClick="return confirm('确定删除此项吗？')">删除</a> 
    </td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
  <tr bgcolor="#FFFFFF"> 
    <td height="35" colspan="5">&nbsp;</td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
</body>
</html>
