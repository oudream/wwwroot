<script language="JavaScript">
<!--
function getsort(targ,selObj,restore){ //v3.0
  eval(targ+".location='subs_view.asp?sort_id="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function getbrand(jsort_id,targ,selObj,restore){ //v3.0
  eval(targ+".location='subs_view.asp?sort_id="+jsort_id+"&brand_id="+selObj.options[selObj.selectedIndex].value+"'");
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
sort_id=request("sort_id")
if sort_id="" then sort_id=0
brand_id=request("brand_id")
if brand_id="" then brand_id=0

Set rs=Server.CreateObject("ADODB.RecordSet")
Set prs=Server.CreateObject("ADODB.RecordSet")
%>
<table width="98%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr align="center"> 
    <td height="30" colspan="7" bgcolor="#FFFFFF"><table align="center" cellpadding=0 cellspacing=0>
        <tbody>
          <tr> 
            <td align="right"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        size=2>品种</font></b></td>
            <td width=114 align="right"> <select name="sort_id" class="display_dropdown" id="sort_id" onChange="getsort('self',this,0)">
          <%
if sort_id=0 then
response.Write("<option value='' selected>请选择……</option>")
else
psql="select * from sort where sort_id=" & sort_id
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
%>
          <option value=<%=prs("sort_id")%> selected><%=prs("sname")%></option>
          <%
end if
prs.close
end if
%>
<%
if sort_id=0 then
%>
                <option value="" selected>全部品种</option>
<%
else
%>
                <option value="">全部品种</option>
<%
end if
%>
<%
psql="select * from sort where sort_id<>" & sort_id
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
          <option value=<%=prs("sort_id")%>><%=prs("sname")%></option>
          <%
prs.movenext
loop
end if
prs.close
%>
              </select> </td>
            <td width=114 align="right"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        size=2>品牌</font></b></td>
            <td width=114 align="right"> <select name="brand_id" class="display_dropdown" id="brand_id" onChange="getbrand(<%=sort_id%>,'self',this,0)">
          <%
if brand_id=0 then
response.Write("<option value='' selected>请选择……</option>")
else
psql="select * from brand where brand_id=" & brand_id
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
%>
          <option value=<%=prs("brand_id")%> selected><%=prs("bname")%></option>
          <%
end if
prs.close
end if
%>
<%
if sort_id=0 then
psql="SELECT * FROM brand where brand_id<>" & brand_id
else
psql="SELECT * FROM brand where brand_id in ( select brand_id from sb where sort_id="&sort_id&" ) and brand_id<>" & brand_id
end if
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
          <option value=<%=prs("brand_id")%>><%=prs("bname")%></option>
          <%
prs.movenext
loop
end if
prs.close
%>
              </select></td>
          </tr>
        </tbody>
      </table></td>
  </tr>
<%
sql="select * from subs order by sort_id,brand_id,code"
if sort_id<>0 and brand_id=0 then sql="select * from subs where sort_id="&sort_id&" order by sort_id,code"
if sort_id=0 and brand_id<>0 then sql="select * from subs where brand_id="&brand_id&" order by brand_id,code"
if sort_id<>0 and brand_id<>0 then sql="select * from subs where  brand_id="&brand_id&" and sort_id=" & sort_id &" order by sort_id,brand_id,code"

rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' None data ! ');</SCRIPT>")
%>
  <tr><td height="22" colspan="8" align="center" class="header"><a href="subs_add.asp" target="mainFrame">商品添加</a></td></tr>
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
    <td height="50" colspan="8" align="center" class="header">
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
                <td width="50" height="30"><a href="subs_view.asp?sort_id=<%=sort_id%>&brand_id=<%=brand_id%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
<%
else
zk=zk+1
%>
              </tr>
              <tr> 
                <td width="50" height="30"><a href="subs_view.asp?sort_id=<%=sort_id%>&brand_id=<%=brand_id%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
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
    <td width="60" height="30" bgcolor="#FFFFFF">种类</td>
    <td width="100" bgcolor="#FFFFFF">品牌</td>
    <td width="130" bgcolor="#FFFFFF">型号</td>
    <td width="180" bgcolor="#FFFFFF">规格参数</td>
    <td width="30" bgcolor="#FFFFFF">单位</td>
    <td width="50" bgcolor="#FFFFFF">进货价</td>
    <td width="50" bgcolor="#FFFFFF">零售价</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr> 
    <td width="60" height="25" bgcolor="#FFFFFF"> <%
ssql="select * from sort where sort_id="&rs("sort_id") 
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.open ssql,conn,1,1 
if not(srs.eof or srs.bof) then response.Write(srs("sname")) 
srs.close
	%> 
      &nbsp;</td>
    <td width="100" bgcolor="#FFFFFF"> <%
ssql="select * from brand where brand_id="&rs("brand_id") 
srs.open ssql,conn,1,1 
if not(srs.eof or srs.bof) then response.Write(srs("bname"))
srs.close
set srs=nothing
	%> &nbsp;</td>
    <td width="130" bgcolor="#FFFFFF"><a href="subs_edit.asp?subs_id=<%=rs("subs_id")%>"><%=rs("code")%></a>&nbsp;</td>
    <td width="180" bgcolor="#FFFFFF"><%=rs("sproperties")%>&nbsp;</td>
    <td width="30" bgcolor="#FFFFFF"><%=rs("suit")%>&nbsp;</td>
    <td width="50" bgcolor="#FFFFFF"><%=rs("pricein")%>&nbsp;</td>
    <td width="50" bgcolor="#FFFFFF"><%=rs("priceout")%>&nbsp;</td>
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
