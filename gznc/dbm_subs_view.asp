<script language="JavaScript">

function checkform()
{
	if(! isNumber(form_administrator.zbig.value)) {
		alert('请用数字填写');
		form_administrator.zbig.focus();
		return false;
		}
	if(! isNumber(form_administrator.zsmall.value)) {
		alert('请用数字填写');
		form_administrator.zsmall.focus();
		return false;
		}
	return true;
}

function isNumber(name)
{
	for(i = 0; i < name.length; i++) {
		if (!((name.charAt(i) >= "0" && name.charAt(i) <= "9") || name.charAt(i) == "."))
			return false;
	}
	return true;
}

function isEnglish(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charCodeAt(i) > 128)
			return false;
	}
	return true;
}
</script>

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<title></title>
<LINK href="dbm_css.css" type=text/css rel=stylesheet>
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
conn.execute "delete * from dbm_subs where subs_id=" & request("subs_id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('删除成功');</SCRIPT>")
end if
%>



<%
zname=trim(request("zname"))
zname=replace(zname,"'","''")
zstate=trim(request("zstate"))
zstate=replace(zstate,"'","''")
zsproperties=trim(request("zsproperties"))
zsproperties=replace(zsproperties,"'","''")

zbig=request("zbig")
zsmall=request("zsmall")

zothers=trim(request("zothers"))
zothers=replace(zothers,"'","''")
%>








        
  
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <form name="form_administrator" id="form_administrator" method="post" action="dbm_subs_view.asp" onSubmit="return checkform();">
    <tr align="center" valign="middle"> 
      <td width="340" height="30">名称 
        <input name="zname" type="text" id="zname" size="20" maxlength="100" value="<%=zname%>"></td>
      <td height="30" colspan="2">其它参数 
        <input name="zothers" type="text" id="zothers" size="20" maxlength="100" value="<%=zothers%>"> 
      </td>
      <td width="112" rowspan="2"><input type="submit" name="Submit" value="查询>>"></td>
    </tr>
  </form>
</table>








<%
Set rs=Server.CreateObject("ADODB.RecordSet")
Set srs=Server.CreateObject("ADODB.RecordSet") 
%>
<br>
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <%
sql="select * from dbm_subs where sname<> ''"

if zname<>"" then sql=sql & " and sname like '%"&zname&"%'"
if zothers<>"" then sql=sql & " and ( subs_id like '%"&zothers&"%' or memo like '%"&zothers&"%')"

sql = sql &" order by subs_id DESC"

rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' 没有合要求的作品 ! ');</SCRIPT>")
%>
  <tr> 
    <td colspan="8" align="center" class="header">&nbsp;</td>
  </tr>
  <%
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
  <tr bgcolor="#FFFFFF"> 
    <td colspan="8" align="center" class="header"><a href="dbm_subs_view.asp"><font color="#FF0000" size="+1"><strong>作品查看</strong></font></a> 
    </td>
  </tr>
  <tr align="center"> 
    <td width="120" height="31" bgcolor="#FFFFFF">作品名称</td>
    <td width="350" bgcolor="#FFFFFF">简介(只显示前100个字符)</td>
    <td width="130" bgcolor="#FFFFFF">操作</td>
  </tr>
<%
do while not rs.eof 
%>
  <tr align="center"> 
    <td width="120" height="30" bgcolor="#FFFFFF"><%=rs("sname")%>&nbsp;</td>
    <td width="350" bgcolor="#FFFFFF">
      <% if len(rs("memo"))>100 then%>
      <%=left(rs("memo"),95)%>.. 
      <%else%>
      <%=rs("memo")%> 
      <%end if%>&nbsp;
	</td>
    <td width="130" bgcolor="#FFFFFF"><a href="dbm_subs_viewall.asp?subs_id=<%=rs("subs_id")%>" target="_blank">详细</a>
      |<a href="dbm_subs_edit.asp?subs_id=<%=rs("subs_id")%>">编辑</a>
      |<a href="dbm_subs_view.asp?subs_id=<%=rs("subs_id")%>&option=del" onClick="return confirm('确定删除此项吗？')">删除</a></td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
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
          <td><a href="dbm_subs_view.asp?zname=<%=zname%>&zstate=<%=zstate%>&zsproperties=<%=zsproperties%>&zbig=<%=zbig%>&zsmall=<%=zsmall%>&zothers=<%=zothers%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
          <%
else
zk=zk+1
%>
        </tr>
        <tr> 
          <td><a href="dbm_subs_view.asp?zname=<%=zname%>&zstate=<%=zstate%>&zsproperties=<%=zsproperties%>&zbig=<%=zbig%>&zsmall=<%=zsmall%>&zothers=<%=zothers%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
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
set srs=nothing
%>
</body>
</html>
