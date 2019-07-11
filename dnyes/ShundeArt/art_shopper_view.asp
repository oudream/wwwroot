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
if trim(request("option"))="del" then
conn.execute "delete * from art_shopper where art_shopper_id=" & request("art_shopper_id") 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('删除成功');</SCRIPT>")
end if
%>






<%
zname=trim(request("zname"))
zname=replace(zname,"'","''")

zaddr=trim(request("zaddr"))
zaddr=replace(zaddr,"'","''")

ztel=trim(request("ztel"))
ztel=replace(ztel,"'","''")

zworkplace=trim(request("zworkplace"))
zworkplace=replace(zworkplace,"'","''")

zduty=request("zduty")
zduty=request("zduty")

zothers=trim(request("zothers"))
zothers=replace(zothers,"'","''")
%>








        
  <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
<form name="form_administrator" id="form_administrator" method="post" action="art_shopper_view.asp">
    <tr align="center" valign="middle"> 
      <td height="30">名称： 
        <input name="zname" type="text" id="zname" size="20" maxlength="100" value="<%=zname%>"></td>
      <td height="30"> 地址： 
        <input name="zaddr" type="text" id="zaddr" size="20" maxlength="100" value="<%=zaddr%>"></td>
      <td> 电话： 
        <input name="ztel" type="text" id="ztel" size="20" maxlength="100" value="<%=ztel%>"></td>
      <td rowspan="2"><input type="submit" name="Submit" value="查询>>"></td>
    </tr>
    <tr align="center" valign="middle"> 
      <td height="30">工作单位： 
        <input name="zworkplace" id="zworkplace" type="text" size="20" maxlength="100" value="<%=zworkplace%>">
      </td>
      <td height="30">职务： 
        <input name="zduty" type="text" id="zduty" size="20" maxlength="100" value="<%=zduty%>"></td>
      <td>其它： 
        <input name="zothers" type="text" id="zothers" size="20" maxlength="100" value="<%=zothers%>"></td>
    </tr>
      </form>
  </table>


<br>
<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from art_shopper where allname<>'' " 

if zname<>"" then sql=sql & " and allname like '%"&zname&"%'"
if zaddr<>"" then sql=sql & " and addr like '%"&zaddr&"%'"
if ztel<>"" then sql=sql & " and tel like '%"&ztel&"%'"
if zworkplace<>"" then sql=sql & " and workplace like '%"&zworkplace&"%'"
if zduty<>"" then sql=sql & " and duty like '%"&zduty&"%'"
if zothers<>"" then sql=sql & " and ( zip like '%"&zothers&"%' or email like '%"&zothers&"%' or sex like '%"&zothers&"%' or nation like '%"&zothers&"%' or original like '%"&zothers&"%' or birth like '%"&zothers&"%' or memo like '%"&zothers&"%')"

sql = sql &" order by art_shopper_id DESC"

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
    <td height="30" colspan="8" align="center" class="header"><font color="#FF0000" size="+1"><strong>个人客户买家列表</strong></font></td>
  </tr>
  <tr align="center"> 
    <td width="100" height="30" bgcolor="#FFFFFF">姓名</td>
    <td bgcolor="#FFFFFF">地址</td>
    <td width="120" bgcolor="#FFFFFF">电话</td>
    <td width="150" bgcolor="#FFFFFF">工作单位</td>
    <td width="60" bgcolor="#FFFFFF">职务</td>
    <td width="150" bgcolor="#FFFFFF">操作</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr align="center"> 
    <td height="25" bgcolor="#FFFFFF"><%=rs("allname")%></td>
    <td bgcolor="#FFFFFF"><%=rs("addr")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("tel")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("workplace")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("duty")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><a href="art_shopper_viewall.asp?art_shopper_id=<%=rs("art_shopper_id")%>" target="_blank">详细</a> | <a href="art_shopper_edit.asp?art_shopper_id=<%=rs("art_shopper_id")%>">编辑</a> | <a href="art_shopper_view.asp?art_shopper_id=<%=rs("art_shopper_id")%>&option=del" onClick="return confirm('确定删除此项吗？')">删除</a></td>
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
          <td><a href="art_shopper_view.asp?zname=<%=zname%>&zaddr=<%=zaddr%>&ztel=<%=ztel%>&zworkplace=<%=zworkplace%>&zcap=<%=zcap%>&zothers=<%=zothers%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
          <%
else
zk=zk+1
%>
        </tr>
        <tr> 
          <td><a href="art_shopper_view.asp?zname=<%=zname%>&zaddr=<%=zaddr%>&ztel=<%=ztel%>&zworkplace=<%=zworkplace%>&zcap=<%=zcap%>&zothers=<%=zothers%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
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