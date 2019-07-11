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
conn.execute "delete * from art_artist_society where art_artist_society_id=" & request("art_artist_society_id")
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

zfax=trim(request("zfax"))
zfax=replace(zfax,"'","''")

zcap=request("zcap")
zcap=request("zcap")

zothers=trim(request("zothers"))
zothers=replace(zothers,"'","''")
%>


        
  <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
<form name="form_administrator" id="form_administrator" method="post" action="art_artist_society_view.asp">
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
      <td height="30">传真： 
        <input name="zfax" id="zfax" type="text" size="20" maxlength="100" value="<%=zfax%>">
      </td>
      <td height="30">会长： 
        <input name="zcap" type="text" id="zcap" size="20" maxlength="100" value="<%=zcap%>"></td>
      <td>其它： 
        <input name="zothers" type="text" id="zothers" size="20" maxlength="100" value="<%=zothers%>"></td>
    </tr>
      </form>
  </table>



<br>
<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from art_artist_society where allname<>'' " 

if zname<>"" then sql=sql & " and allname like '%"&zname&"%'"
if zaddr<>"" then sql=sql & " and addr like '%"&zaddr&"%'"
if ztel<>"" then sql=sql & " and tel like '%"&ztel&"%'"
if zfax<>"" then sql=sql & " and fax like '%"&zfax&"%'"
if zcap<>"" then sql=sql & " and cap like '%"&zcap&"%'"
if zothers<>"" then sql=sql & " and ( zip like '%"&zothers&"%' or email like '%"&zothers&"%' or web_addr like '%"&zothers&"%' or creation_date like '%"&zothers&"%' or brief like '%"&zothers&"%' or memo like '%"&zothers&"%')"

sql = sql &" order by art_artist_society_id DESC"

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
    <td height="30" colspan="8" align="center" class="header"><font color="#FF0000" size="+1"><strong>美术家协会列表</strong></font></td>
  </tr>
  <tr align="center"> 
    <td width="80" height="30" bgcolor="#FFFFFF">全称</td>
    <td width="60" bgcolor="#FFFFFF">地址</td>
    <td width="100" bgcolor="#FFFFFF">电话</td>
    <td width="150" bgcolor="#FFFFFF">传真</td>
    <td width="80" bgcolor="#FFFFFF">会长</td>
    <td width="150" bgcolor="#FFFFFF">操作</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr align="center"> 
    <td height="25" bgcolor="#FFFFFF"><%=rs("allname")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("addr")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("tel")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("fax")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("cap")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><a href="art_artist_society_viewall.asp?art_artist_society_id=<%=rs("art_artist_society_id")%>" target="_blank">详细</a>&nbsp; | &nbsp;<a href="art_artist_society_edit.asp?art_artist_society_id=<%=rs("art_artist_society_id")%>">编辑</a>&nbsp; | &nbsp;<a href="art_artist_society_view.asp?art_artist_society_id=<%=rs("art_artist_society_id")%>&option=del" onClick="return confirm('确定删除此项吗？')">删除</a>&nbsp;</td>
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
          <td><a href="art_artist_society_view.asp?zname=<%=zname%>&zaddr=<%=zaddr%>&ztel=<%=ztel%>&zfax=<%=zfax%>&zcap=<%=zcap%>&zothers=<%=zothers%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
          <%
else
zk=zk+1
%>
        </tr>
        <tr> 
          <td><a href="art_artist_society_view.asp?zname=<%=zname%>&zaddr=<%=zaddr%>&ztel=<%=ztel%>&zfax=<%=zfax%>&zcap=<%=zcap%>&zothers=<%=zothers%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
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
