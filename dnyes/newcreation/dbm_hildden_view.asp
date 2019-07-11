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

<!--#include file="dbm_adminconn.asp" -->



<%
if trim(request("option"))="del" then
conn.execute "delete * from dbm_hildden where dbm_hildden_id=" & request("dbm_hildden_id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('删除成功');</SCRIPT>")
end if
%>


<br>
<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from dbm_hildden order by dbm_hildden_id DESC"
rs.open sql,conn,1,1 
if rs.eof or rs.bof then 
%>
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="20" align="center">暂无您所要的数据!</td>
  </tr>
</table>
<%
response.End()
end if
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=15
rs.AbsolutePage=pagecount
%>

<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr bgcolor="#FFFFFF"> 
    <td height="30" colspan="8" align="center" class="header"><font color="#FF0000" size="+1"><strong>收藏夹查看修改</strong></font></td>
  </tr>
  <tr align="center"> 
    <td width="120" height="31" bgcolor="#FFFFFF">标题</td>
    <td width="200" bgcolor="#FFFFFF">网址</td>
    <td width="256" bgcolor="#FFFFFF">简介</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr align="center"> 
    <td height="25" bgcolor="#FFFFFF"><%=rs("htitle")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><a href="http://<%=rs("web_addr")%>" target="_blank"><%=rs("web_addr")%></a>&nbsp;</td>
    <td bgcolor="#FFFFFF"><a href="dbm_hildden_viewall.asp?dbm_hildden_id=<%=rs("dbm_hildden_id")%>">详细</a>
      |<a href="dbm_hildden_edit.asp?dbm_hildden_id=<%=rs("dbm_hildden_id")%>">编辑</a>
      |<a href="dbm_hildden_view.asp?dbm_hildden_id=<%=rs("dbm_hildden_id")%>&option=del" onClick="return confirm('确定删除此项吗？')">删除</a></td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="70%" height="30"> <table border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <%
zj=1
zk=1
for zi=1 to rs.pagecount
if zj<=zk*15 then
%>
          <td><a href="dbm_hildden_view.asp?page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
          <%
else
zk=zk+1
%>
        </tr>
        <tr> 
          <td><a href="dbm_hildden_view.asp?page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
          <%
end if
zj=zj+1
next
%>
        </tr>
      </table></td>
    <td width="33%" align="left"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="right"> Record :<font color=red><b><%=rs.recordcount%></b></font> 
            Page <font color=red><b><%=pagecount%></b></font> of <font color=red><b><%=rs.pagecount%></b></font> 
          </td>
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
