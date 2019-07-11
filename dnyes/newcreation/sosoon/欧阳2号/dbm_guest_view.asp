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
    <td height="20">&nbsp;</td>
  </tr>
</table>

<!--#include file="dbm_adminconn.asp" -->





<%
if trim(request("option"))="del" then
conn.execute "delete * from dbm_guest where guest_id=" & request("guest_id")

  If not Err Then
  	zrootpath = Server.MapPath(".")
    Response.Redirect("dbm_folder_option.asp?Action=DelFolder&FName="+zrootpath+"\"+request("allname")+"&zurl=dbm_guest_view.asp")
  End If		

response.Write("<SCRIPT LANGUAGE=JavaScript> alert('删除成功');</SCRIPT>")
end if
%>




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

        
  
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <form name="form_administrator" id="form_administrator" method="post" action="dbm_guest_view.asp" onSubmit="return checkform();">
    <tr align="center" valign="middle"> 
      <td width="227" height="30">名称 </td>
      <td width="222" height="30">地址 </td>
      <td width="281">其它 </td>
      <td width="81">&nbsp;</td>
    </tr>
    <tr align="center" valign="middle"> 
      <td height="30"><input name="zname" type="text" id="zname2" size="20" maxlength="100" value="<%=zname%>"></td>
      <td height="30"><input name="zaddr" type="text" id="zaddr2" size="20" maxlength="100" value="<%=zaddr%>"></td>
      <td><input name="zothers" type="text" id="zothers2" size="20" maxlength="100" value="<%=zothers%>"></td>
      <td><input type="submit" name="Submit" value="查询>>"></td>
    </tr>
  </form>
</table>













<br>
<%
sql="select * from dbm_guest where allname<> ''"

if zname<>"" then sql=sql & " and allname like '%"&zname&"%'"
if zaddr<>"" then sql=sql & " and addr like '%"&zaddr&"%'"
if zothers<>"" then sql=sql & " and ( tel like '%"&zothers&"%' or fax like '%"&zothers&"%' or zip like '%"&zothers&"%' or email like '%"&zothers&"%' or gid like '%"&zothers&"%' or pwd like '%"&zothers&"%' or memo like '%"&zothers&"%')"

sql=sql&" order by allname"

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
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr bgcolor="#FFFFFF"> 
    <td height="30" colspan="8" align="center" class="header"><font color="#FF0000" size="+1"><strong>用户列表</strong></font> 
    </td>
  </tr>
  <tr align="center"> 
    <td width="18%" height="31" bgcolor="#FFFFFF">名称(公司或个人)</td>
    <td width="23%" bgcolor="#FFFFFF">地址</td>
    <td width="15%" bgcolor="#FFFFFF">电话</td>
    <td width="15%" bgcolor="#FFFFFF">账号</td>
    <td width="15%" bgcolor="#FFFFFF">密码</td>
    <td width="14%" bgcolor="#FFFFFF">操作</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr align="center"> 
    <td width="18%" height="25" bgcolor="#FFFFFF"><%=rs("allname")%></td>
    <td width="23%" bgcolor="#FFFFFF"><%=rs("addr")%>&nbsp;</td>
    <td width="15%" bgcolor="#FFFFFF"><%=rs("tel")%>&nbsp;</td>
    <td width="15%" bgcolor="#FFFFFF"><%=rs("gid")%>&nbsp;</td>
    <td width="15%" bgcolor="#FFFFFF"><%=rs("pwd")%>&nbsp;</td>
    <td width="14%" bgcolor="#FFFFFF"><a href="dbm_guest_edit.asp?guest_id=<%=rs("guest_id")%>">编辑</a> 
      | <a href="dbm_guest_view.asp?guest_id=<%=rs("guest_id")%>&allname=<%=rs("allname")%>&option=del" onClick="return confirm('确定删除此项吗？\n \n将删除客户、及其文件夹、以及文件夹内所有东西')">删除</a> 
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
    <td height="35" colspan="8">&nbsp;</td>
  </tr>
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
          <td><a href="dbm_guest_view.asp?page=<%=zj%>&zname=<%=zname%>&zsex=<%=zsex%>&zaddr=<%=zaddr%>&zothers=<%=zothers%>">&nbsp;<%=zj%>&nbsp;</a></td>
          <%
else
zk=zk+1
%>
        </tr>
        <tr> 
          <td><a href="dbm_guest_view.asp?page=<%=zj%>&zname=<%=zname%>&zsex=<%=zsex%>&zaddr=<%=zaddr%>&zothers=<%=zothers%>">&nbsp;<%=zj%>&nbsp;</a></td>
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
