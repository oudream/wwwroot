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
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="30" colspan="9" class="header"><font color="#FF0000"><strong>管理员留言</strong></font></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="22" colspan="9" class="header">&nbsp; </td>
  </tr>
  <tr> 
    <td align="center" bgcolor="#FFFFFF">主题</td>
    <td width="100" align="center" bgcolor="#FFFFFF">留言人</td>
    <td width="150" align="center" bgcolor="#FFFFFF">留言时间</td>
    <td align="center" bgcolor="#FFFFFF">留言内容</td>
    <td width="150" align="center" bgcolor="#FFFFFF">操作</td>
  </tr>
  <%
Set urs=Server.CreateObject("ADODB.RecordSet") 
do while not rs.eof 
%>
  <tr> 
    <td align="center" bgcolor="#FFFFFF"><%=rs("subject")%></td>
    <td align="center" bgcolor="#FFFFFF"> <%
usql="select * from art_user where id="&rs("user_id") 
Set urs=Server.CreateObject("ADODB.RecordSet") 
urs.open usql,conn,1,1 
uname=urs("uname")
urs.close
%> 
      <%=uname%>&nbsp; </td>
    <td align="center" bgcolor="#FFFFFF"><%=rs("inserttime")%>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"><%=rs("memo")%>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"><a href="art_message_all.asp????">查看</a> 
      | <a href="art_message_edit.asp?zid=<%=rs("id")%>">编辑</a> | <a href="art_message_del.asp?????">删除</a></td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
set urs=nothing
%>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="35" colspan="9"><font size="4"><strong><a href="art_message_add.asp">留言添加</a></strong></font></td>
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
          <td><a href="art_message_view.asp?page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
          <%
zj=zj+1
next
%>
        </tr>
      </table></td>
    <td width="26%" align="left"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td> <div align="center"> Record : <font color=red><b><%=rs.recordcount%></b></font> 
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
