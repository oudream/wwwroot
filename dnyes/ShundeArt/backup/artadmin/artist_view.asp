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
zname=trim(request("zname"))
zname=replace(zname,"'","''")
zaddr=trim(request("zaddr"))
zaddr=replace(zaddr,"'","''")
zsex=trim(request("zsex"))
zsex=replace(zsex,"'","''")
zothers=trim(request("zothers"))
zothers=replace(zothers,"'","''")
%>







<form name="form_administrator" id="form_administrator" method="post" action="artist_view.asp" onSubmit="return checkform();">
        
  <table width="97%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
    <tr align="center" valign="middle"> 
      <td width="227" height="30">名称 </td>
      <td width="222" height="30">地址 </td>
      <td width="168">性别 </td>
      <td width="281">其它 </td>
      <td width="81">&nbsp;</td>
    </tr>
    <tr align="center" valign="middle">
      <td height="30"><input name="zname" type="text" id="zname2" size="20" maxlength="100" value="<%=zname%>"></td>
      <td height="30"><input name="zaddr" type="text" id="zaddr2" size="20" maxlength="100" value="<%=zaddr%>"></td>
      <td><select name="zsex" id="zsex">
          <%
if zsex="" then
%>
          <option value="男" selected>男</option>
          <option value="女">女</option>
          <%
end if
%>
          <%
if zsex="男" then
%>
          <option value="男" selected>男</option>
          <option value="女">女</option>
          <%
end if
%>
          <%
if zsex="女" then
%>
          <option value="男">男</option>
          <option value="女" selected>女</option>
          <%
end if
%>
        </select></td>
      <td><input name="zothers" type="text" id="zothers2" size="20" maxlength="100" value="<%=zothers%>"></td>
      <td><input type="submit" name="Submit" value="查询>>"></td>
    </tr>
  </table>
      </form> 













<%
sql="select * from art_artist where allname<> ''"

if zname<>"" then sql=sql & " and allname like '%"&zname&"%'"
if zaddr<>"" then sql=sql & " and addr like '%"&zaddr&"%'"
if zsex<>"" then sql=sql & " and sex like '%"&zsex&"%'"
if zothers<>"" then sql=sql & " and ( simname like '%"&zothers&"%' or nation like '%"&zothers&"%' or native like '%"&zothers&"%' or born_time like '%"&zothers&"%' or graduation like '%"&zothers&"%' or special like '%"&zothers&"%' or job_now like '%"&zothers&"%' or job_old like '%"&zothers&"%' or job_rank like '%"&zothers&"%' or expert like '%"&zothers&"%' or awarded like '%"&zothers&"%' or standest like '%"&zothers&"%' or interest like '%"&zothers&"%' or tel like '%"&zothers&"%' or fax like '%"&zothers&"%' or zip like '%"&zothers&"%' or email like '%"&zothers&"%' or memo like '%"&zothers&"%')"

sql=sql&" order by allname DESC"

Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' 你好，没有你要查询的数据。多谢合作，美协 ! ');</SCRIPT>")
response.End()
end if
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=20
rs.AbsolutePage=pagecount
%>
<table width="97%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr bgcolor="#FFFFFF"> 
    <td height="22" colspan="8" align="center" class="header">
	
	
	
	
	
	
	
	
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="70%"> <table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
<%
zj=1
zk=1
for zi=1 to rs.pagecount
if zj<=zk*15 then
%>
                <td width="50" height="30"><a href="artist_view.asp?page=<%=zj%>&zname=<%=zname%>&zsex=<%=zsex%>&zaddr=<%=zaddr%>&zothers=<%=zothers%>">&nbsp;<%=zj%>&nbsp;</a></td>
<%
else
zk=zk+1
%>
              </tr>
              <tr> 
                <td width="50" height="30"><a href="artist_view.asp?page=<%=zj%>&zname=<%=zname%>&zsex=<%=zsex%>&zaddr=<%=zaddr%>&zothers=<%=zothers%>">&nbsp;<%=zj%>&nbsp;</a></td>
<%
end if
zj=zj+1
next
%>
              </tr>
            </table></td>
          <td width="33%" align="left"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td align="right"> Record :<font color=red><b><%=rs.recordcount%></b></font> Page <font color=red><b><%=pagecount%></b></font> of <font color=red><b><%=rs.pagecount%></b></font> </td>
              </tr>
            </table></td>
        </tr>
      </table>	
	
	
	
	
	
	
	
	
	
	
	 </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="22" colspan="8" align="center" class="header">客户列表</td>
  </tr>
  <tr> 
    <td width="80" height="30" bgcolor="#FFFFFF">全称</td>
    <td width="60" bgcolor="#FFFFFF">雅号</td>
    <td width="100" bgcolor="#FFFFFF">性别</td>
    <td width="150" bgcolor="#FFFFFF">地址</td>
    <td width="80" bgcolor="#FFFFFF">出生年月</td>
    <td width="150" bgcolor="#FFFFFF">备注</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr> 
    <td height="25" bgcolor="#FFFFFF"><%=rs("allname")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><a href="artist_edit.asp?artist_id=<%=rs("artist_id")%>"><%=rs("simname")%></a>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("sex")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("addr")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("born_time")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("memo")%>&nbsp;</td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
  <tr bgcolor="#FFFFFF"> 
    <td height="35" colspan="8">&nbsp;</td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
</body>
</html>
