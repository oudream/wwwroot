<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<title></title>
</head>
<body topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>

<!--#include file="dbm_adminconn.asp" -->




<%
if trim(request("option"))="del" then
conn.execute "delete * from dbm_note where id=" & request("zid")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('删除成功');</SCRIPT>")
end if
%>





<%
if request("option_id_del")="succ" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('删除成功');</SCRIPT>")
end if

sql="select * from dbm_note order by id desc" 
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
    <td height="22" colspan="9"><font color="#FF0000"><strong>留言列表</strong></font></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="22" colspan="9" class="header">&nbsp; </td>
  </tr>
  <tr> 
    <td align="center" bgcolor="#FFFFFF">题摘</td>
    <td width="120" align="center" bgcolor="#FFFFFF">留言用户</td>
    <td width="150" align="center" bgcolor="#FFFFFF">留言时间</td>
    <td align="center" bgcolor="#FFFFFF">留言</td>
    <td width="150" align="center" bgcolor="#FFFFFF">操作</td>
  </tr>
  <%
Set urs=Server.CreateObject("ADODB.RecordSet") 
do while not rs.eof 
%>
  <tr> 
    <td align="center" bgcolor="#FFFFFF"><%=rs("subject")%></td>
    <td align="center" bgcolor="#FFFFFF"> <%
usql="select * from dbm_user where user_id="&rs("user_id") 
Set urs=Server.CreateObject("ADODB.RecordSet") 
urs.open usql,conn,1,1 
allname=urs("allname")
urs.close
%> 
      <%=allname%>&nbsp; </td>
    <td align="center" bgcolor="#FFFFFF"><%=rs("inserttime")%>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"><%=rs("memo")%>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"> <a href="dbm_note_edit.asp?zid=<%=rs("id")%>">编辑</a> | <a href="dbm_note_view.asp?zid=<%=rs("id")%>&option=del"  onClick="return confirm('确定删除此项吗？\n \n 将删除此留言')">删除</a></td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
set urs=nothing
%>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="35" colspan="9">&nbsp;</td>
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
          <td><a href="dbm_note_view.asp?page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
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
