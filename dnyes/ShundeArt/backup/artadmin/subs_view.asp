<script language="JavaScript">
<!--
function getartist(targ,selObj,restore){ //v3.0
  eval(targ+".location='subs_view.asp?artist_id="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
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
zstate=trim(request("zstate"))
zstate=replace(zstate,"'","''")
zsproperties=trim(request("zsproperties"))
zsproperties=replace(zsproperties,"'","''")

zbig=request("zbig")
zsmall=request("zsmall")

zothers=trim(request("zothers"))
zothers=replace(zothers,"'","''")
%>








<form name="form_administrator" id="form_administrator" method="post" action="subs_view.asp" onSubmit="return checkform();">
        
  <table width="97%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
    <tr align="center" valign="middle"> 
      <td width="243" height="30">名称 
        <input name="zname" type="text" id="zname" size="20" maxlength="100" value="<%=zname%>"></td>
      <td width="139" height="30">状态 
		  <select name="zstate" id="zstate">
          <%
if zstate="" then
%>
            <option value="已出售">已出售</option>
            <option value="未出售" selected>未出售</option>
            <option value="收藏">收藏</option>
          <%
end if
%>
          <%
if zstate="已出售" then
%>
            <option value="已出售" selected>已出售</option>
            <option value="未出售">未出售</option>
            <option value="收藏">收藏</option>
          <%
end if
%>
          <%
if zstate="未出售" then
%>
            <option value="已出售">已出售</option>
            <option value="未出售" selected>未出售</option>
            <option value="收藏">收藏</option>
          <%
end if
%>
          <%
if zstate="收藏" then
%>
            <option value="已出售">已出售</option>
            <option value="未出售">未出售</option>
            <option value="收藏" selected>收藏</option>
          <%
end if
%>
        </select></td>
      <td width="258"> 规格参数 
        <input name="zsproperties" type="text" id="zsproperties" size="20" maxlength="100" value="<%=zsproperties%>"></td>
      <td width="112" rowspan="2"><input type="submit" name="Submit" value="查询>>"></td>
    </tr>
    <tr align="center" valign="middle"> 
      <td height="30" colspan="2">作品价格大于 
        <input name="zbig" id="zbig" type="text" size="10" maxlength="8" value="<%=zbig%>">
        作品价格小于 
        <input name="zsmall" id="zsmall" type="text" size="10" maxlength="8" value="<%=zsmall%>"></td>
      <td>其它参数 
        <input name="zothers" type="text" id="zothers" size="20" maxlength="100" value="<%=zothers%>"></td>
    </tr>
  </table>
      </form>








<%
Set rs=Server.CreateObject("ADODB.RecordSet")
Set prs=Server.CreateObject("ADODB.RecordSet")
%>

<table width="98%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <%
sql="select * from art_subs where sname<> ''"

if zname<>"" then sql=sql & " and sname like '%"&zname&"%'"
if zsproperties<>"" then sql=sql & " and sproperties like '%"&zsproperties&"%'"
if zstate<>"" then sql=sql & " and sstate like '%"&zstate&"%'"
if zbig<>"" then sql=sql & " and price >= " & zbig
if zsmall<>"" then sql=sql & " and price <= " & zsmall
if zothers<>"" then sql=sql & " and ( subs_id like '%"&zothers&"%' or simg like '%"&zothers&"%' or memo like '%"&zothers&"%')"

sql = sql &" order by sname DESC"

rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' 此作者还没有任何作品 ! ');</SCRIPT>")
%>
  <tr> 
    <td height="22" colspan="8" align="center" class="header"><a href="subs_add.asp" target="mainFrame">作品查看</a></td>
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
    <td height="50" colspan="8" align="center" class="header"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="70%"> <table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
zk=1
for zi=1 to rs.pagecount
if zj<=zk*15 then
%>
                <td width="50" height="30"><a href="subs_view.asp?zname=<%=zname%>&zstate=<%=zstate%>&zsproperties=<%=zsproperties%>&zbig=<%=zbig%>&zsmall=<%=zsmall%>&zothers=<%=zothers%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
                <%
else
zk=zk+1
%>
              </tr>
              <tr> 
                <td width="50" height="30"><a href="subs_view.asp?zname=<%=zname%>&zstate=<%=zstate%>&zsproperties=<%=zsproperties%>&zbig=<%=zbig%>&zsmall=<%=zsmall%>&zothers=<%=zothers%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
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
      </table></td>
  </tr>
  <tr> 
    <td width="80" height="30" bgcolor="#FFFFFF">作品名称</td>
    <td width="60" height="30" bgcolor="#FFFFFF">作者</td>
    <td width="100" bgcolor="#FFFFFF">规格参数</td>
    <td width="50" bgcolor="#FFFFFF">市场价</td>
    <td width="50" bgcolor="#FFFFFF">状态</td>
    <td width="200" bgcolor="#FFFFFF">备注</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr> 
    <td width="80" bgcolor="#FFFFFF"><a href="subs_edit.asp?subs_id=<%=rs("subs_id")%>"><%=rs("sname")%>&nbsp;</a></td>
    <td width="60" height="25" bgcolor="#FFFFFF"> <%
ssql="select * from art_subs where artist_id="&rs("artist_id") 
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.open ssql,conn,1,1 
if not(srs.eof or srs.bof) then response.Write(srs("sname")) 
srs.close
	%> &nbsp;</td>
    <td width="100" bgcolor="#FFFFFF"><%=rs("sproperties")%>&nbsp;</td>
    <td width="50" bgcolor="#FFFFFF"><%=rs("price")%>&nbsp;</td>
    <td width="50" bgcolor="#FFFFFF"><%=rs("sstate")%>&nbsp;</td>
    <td width="200" bgcolor="#FFFFFF"><%=rs("memo")%>&nbsp;</td>
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
set prs=nothing
%>
</body>
</html>
