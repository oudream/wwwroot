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
artist_id=request("artist_id")
if artist_id="" then artist_id=0

Set rs=Server.CreateObject("ADODB.RecordSet")
Set prs=Server.CreateObject("ADODB.RecordSet")
%>
<table width="98%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr align="center"> 
    <td height="30" colspan="7" bgcolor="#FFFFFF"><table align="center" cellpadding=0 cellspacing=0>
        <tbody>
          <tr> 
            <td width="53" align="right"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        size=2>作者</font></b></td>
            <td width=140 align="right">
			 <select name="artist_id" class="display_dropdown" id="artist_id" onChange="getartist('self',this,0)">
                <%
if artist_id=0 then
response.Write("<option value='' selected>请选择……</option>")
else
psql="select * from art_artist where artist_id=" & artist_id
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
%>
                <option value=<%=prs("artist_id")%> selected><%=prs("allname")%></option>
                <%
end if
prs.close
end if
%>
                <%
if artist_id=0 then
%>
                <option value="" selected>全部作者</option>
                <%
else
%>
                <option value="">全部作者</option>
                <%
end if
%>
                <%
psql="select * from art_artist where artist_id<>" & artist_id
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
                <option value=<%=prs("artist_id")%>><%=prs("allname")%></option>
                <%
prs.movenext
loop
end if
prs.close
%>
              </select> </td>
          </tr>
        </tbody>
      </table></td>
  </tr>
  <%
sql="select * from art_subs order by artist_id"
if artist_id<>0 then sql="select * from art_subs where artist_id="&artist_id&" order by artist_id"
if artist_id=0 then sql="select * from art_subs where artist_id<>0 order by artist_id"

rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' 此作者还没有任何作品 ! ');</SCRIPT>")
%>
  <tr> 
    <td height="22" colspan="8" align="center" class="header"><a href="subs_add.asp" target="mainFrame">作品添加</a></td>
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
                <td width="50" height="30"><a href="subs_view.asp?artist_id=<%=artist_id%>&brand_id=<%=brand_id%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
                <%
else
zk=zk+1
%>
              </tr>
              <tr> 
                <td width="50" height="30"><a href="subs_view.asp?artist_id=<%=artist_id%>&brand_id=<%=brand_id%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
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
ssql="select * from art_artist where artist_id="&rs("artist_id") 
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.open ssql,conn,1,1 
if not(srs.eof or srs.bof) then response.Write(srs("allname")) 
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
