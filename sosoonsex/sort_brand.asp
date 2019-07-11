<script language="JavaScript">
<!--
function gettarg(targ,selObj,restore){ //v3.0
  eval(targ+".location='sort_brand.asp?sort_id="+selObj.options[selObj.selectedIndex].value+"'");
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
<!--#include file="adminconn.asp" -->
<%
if trim(request("option"))="del" then
conn.execute "delete * from sb where sort_id=" & request("sort_id") & " and brand_id =" & request("brand_id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('成功从该种类中减去');</SCRIPT>")
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
sname=crs("sname")
crs.close
set crs=nothing
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="10">&nbsp;</td>
        </tr>
        <tr> 
          <td height="30" align="center" valign="middle" class="header">种类 -- <font color="#FF0000"><%=sname%></font> -- 所包含品牌的编缉</td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td> 
<!-- ========================================================================================================		  

													cup table ============star

 ======================================================================================================== -->
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td> 
<!-- ========================================================================================================		  

													fixture or result ============star

 ======================================================================================================== -->
            <!-- --------------------------------------------------------------------------------------------------
													assign to "sort_id"
---------------------------------------------------------------------------------------------------- -->
<%
dim abrand()
tsql="select brand_id from sb where sort_id="&sort_id
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if not ( trs.eof or trs.bof ) then
redim abrand(trs.recordcount-1)
for i=0 to ubound(abrand)
abrand(i)=trs("brand_id")
trs.movenext
next
trs.close
set trs=nothing
else
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('该种类没有包含任何品牌');</SCRIPT>")
%>
            
            <table width="96%" border="0" align="center" cellpadding="2" cellspacing="2">
              <tr>
                <td align="center"><strong><a href="sort_brand_select.asp?sort_id=<%=sort_id%>">增添品牌</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="#" onClick="history.go( -1 );">Go Back 
                 </a></strong></td>
              </tr>
            </table> 
<%
response.End()
end if
%>
<!-- ========================================================================================================		  

													fixture or result ============stop

 ======================================================================================================== -->
<!-- --------------------------------------------------------------------------------------------------
													assign to "brand"
---------------------------------------------------------------------------------------------------- -->
            <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
              <tr> 
                <td height="20" colspan="4" align="center" bgcolor="#FFFFFF"><table cellpadding=0 cellspacing=0>
                    <tbody>
                      <tr> 
                        <td align="right"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        size=2>选择你想编缉的种类</font></b></td>
                        <td width=114 align="right"> 
						<select name="sort_id" class="display_dropdown" id="sort_id" onchange="gettarg('self',this,0)">
                            <option selected value="">SELECT ... </option>
                            <%
sql="select * from sort order by sort_id" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
                            <option value=<%=rs("sort_id")%>><%=rs("sname")%> </option>
                            <%
rs.movenext
loop
rs.close
set rs=nothing
%>
                          </select> </td>
                      </tr>
                    </tbody>
                  </table></td>
              </tr>
              <tr> 
                <td width="150" height="30" bgcolor="#FFFFFF">品牌名称</td>
                <td width="350" bgcolor="#FFFFFF">备注</td>
                <td width="100" bgcolor="#FFFFFF">操作</td>
              </tr>
<%
tsql="select * from brand order by bname"
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
%>
              <%
if trs.eof or trs.bof then
response.Write("<tr><td rowspan=5 align=center bgcolor=#FFFFFF>NO DATA</td></tr></table>")
response.End()
trs.close
set trs=nothing
else
for j=0 to trs.recordcount-1
for i=0 to ubound(abrand)
if abrand(i)=trs("brand_id") then
%>
              <tr> 
                <td height="25" bgcolor="#FFFFFF"><%=trs("bname")%></td>
                <td bgcolor="#FFFFFF"> <%=trs("memo")%>&nbsp;</td>
                <td bgcolor="#FFFFFF"><a href="sort_brand.asp?option=del&sort_id=<%=sort_id%>&brand_id=<%=trs("brand_id")%>" onClick="return confirm('是否确认从该种类中减去此品牌')">从该种类中减去</a></td>
              </tr>
              <%
end if
next
trs.movenext
next
end if
%>
              <tr> 
                <td height="40" colspan="4" align="center" bgcolor="#FFFFFF"><strong><a href="sort_brand_select.asp?sort_id=<%=sort_id%>">增添品牌</a></strong></td>
              </tr>
              <%
trs.close
set trs=nothing
%>
            </table></td>
        </tr>
      </table> 
      <!-- ========================================================================================================		  

													cup table ============stop

 ======================================================================================================== -->
    </td>
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
