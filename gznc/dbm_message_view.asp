<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<title></title>
</head>
<LINK href="dbm_css.css" type=text/css rel=stylesheet>
<body topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>

<!--#include file="dbm_adminconn.asp" -->






<%
if trim(request("option"))="del" then
conn.execute "delete * from dbm_message where id=" & request("zid")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('删除成功');</SCRIPT>")
end if
%>



<%
if session("user_adminlevel")=1 then
sql="select * from dbm_message order by id desc" 
else
if session("user_adminlevel")=9 then
sql="select * from dbm_message where guest_id="&session("user_id")&" order by id desc"
else
sql="select * from dbm_message where user_id="&session("user_id")&" order by id desc"
end if
end if
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('没有任何可以编辑的短信');</SCRIPT>")
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
    <td height="30" colspan="9" class="header"><font color="#FF0000"><strong>短信息查看修改</strong></font></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="22" colspan="9" class="header">&nbsp; </td>
  </tr>
  <tr> 
    <td align="center" bgcolor="#FFFFFF">主题</td>
    <td width="100" align="center" bgcolor="#FFFFFF">发短信息人</td>
    <td width="100" align="center" bgcolor="#FFFFFF">收短信息人</td>
    <td width="150" align="center" bgcolor="#FFFFFF">短信息时间</td>
    <td align="center" bgcolor="#FFFFFF">短信息内容</td>
    <td width="150" align="center" bgcolor="#FFFFFF">操作</td>
  </tr>
  <%
Set urs=Server.CreateObject("ADODB.RecordSet") 
do while not rs.eof 
%>
  <tr> 
    <td align="center" bgcolor="#FFFFFF"><%=rs("subject")%></td>












<%
if rs("user_id")=0 then
%>

    <td align="center" bgcolor="#FFFFFF">
<%
usql="select * from dbm_guest where guest_id="&rs("guest_id") 
Set urs=Server.CreateObject("ADODB.RecordSet") 
urs.open usql,conn,1,1 
allname=urs("allname")
urs.close
%> 
      <%=allname%>&nbsp; </td>
    <td align="center" bgcolor="#FFFFFF">新筑</td>

<%
else
%>	  
	  
    <td align="center" bgcolor="#FFFFFF">
<%
usql="select * from dbm_user where user_id="&rs("user_id") 
Set urs=Server.CreateObject("ADODB.RecordSet") 
urs.open usql,conn,1,1 
allname=urs("allname")
urs.close
%> 
      <%=allname%>&nbsp; </td>
    <td align="center" bgcolor="#FFFFFF">
<%
usql="select * from dbm_guest where guest_id="&rs("guest_id") 
Set urs=Server.CreateObject("ADODB.RecordSet") 
urs.open usql,conn,1,1 
allname=urs("allname")
urs.close
%> 
      <%=allname%>&nbsp; </td>

<%
end if
%>	  
	  
	  
	  
	  
	  
	  
	  
	  
    <td align="center" bgcolor="#FFFFFF"><%=rs("inserttime")%>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"><%=rs("memo")%>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"><a href="dbm_message_edit.asp?zid=<%=rs("id")%>">编辑</a> | <a href="dbm_message_view.asp?zid=<%=rs("id")%>&option=del" onClick="return confirm('确定删除此项吗？\n \n 将删除此短信息')">删除</a></td>
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
          <td><a href="dbm_message_view.asp?page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
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
