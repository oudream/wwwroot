<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>

<body>
<!--#include file="adminconn.asp" -->
<%
if trim(request("option"))="add" then
conn.execute "insert into sb (sort_id,brand_id,inserttime) values ("&request("sort_id")&","&request("brand_id")&",'"&now()&"')"
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('增添成功');</SCRIPT>")
end if
%>
<%
sort_id=request("sort_id")
if sort_id="" then sort_id=52
csql="select * from sort where sort_id="&sort_id
Set crs=Server.CreateObject("ADODB.RecordSet")
crs.Open csql,conn,1,1
if crs.eof or crs.bof then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('没有所需要的种类');</SCRIPT>")
response.End()
crs.close
set crs=nothing
end if
cname=crs("sname")
crs.close
set crs=nothing
%>
<%
dim abrand()
tsql="select brand_id from sb where sort_id="&sort_id
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if not ( trs.eof or trs.bof ) then
abrandft="yes"
redim abrand(trs.recordcount-1)
for i=0 to ubound(abrand)
abrand(i)=trs("brand_id")
trs.movenext
next
trs.close
set trs=nothing
end if
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="10">&nbsp;</td>
        </tr>
        <tr> 
          <td height="30" align="center" valign="middle"><strong>在 -- <font color="#FF0000"><%=cname%></font> -- 增添品牌</strong></td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
        <tr> 
          <td width="80%" height="50">品牌名称</td>
          <td width="20%">选项</td>
        </tr>
        <%
tsql="select * from brand where brand_id <> 0"
if abrandft="yes" then
for i=0 to ubound(abrand)
tsql=tsql&" and brand_id <>"&abrand(i)
next
end if
tsql =tsql &" order by bname "
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if not ( trs.eof or trs.bof ) then 
do while not trs.eof
%>
        <tr> 
          <td width="80%" height="25"><%=trs("bname")%>&nbsp;</td>
          <td width="20%"><a href="sort_brand_select.asp?sort_id=<%=sort_id%>&brand_id=<%=trs("brand_id")%>&option=add">增添该品牌</a></td>
        </tr>
        <%
trs.movenext
loop
end if
trs.close
set trs=nothing
%>
        <tr> 
          <td height="38" colspan="2">&nbsp;</td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
