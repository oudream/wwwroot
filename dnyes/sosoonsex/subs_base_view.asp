<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="8" topmargin="8" marginwidth="0" marginheight="0">
<table width="600" height="30" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
</table>

<!--#include file="adminconn.asp" -->
  <table width="700" height="100" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td><iframe id=IFrame src="subs_base_view_in.asp" frameborder=0 width=98% scrolling=no height="100"></iframe></td>
    </tr>
  </table>
<%
pmonth=cint(request("pmonth"))
if pmonth=0 then pmonth=12
pday=cint(request("pday"))
if pday=0 then pday=31

sort_id=request("sort_id")
if sort_id="" then sort_id=0
brand_id=request("brand_id")
if brand_id="" then brand_id=0
subs_id=request("subs_id")
if subs_id="" then subs_id=0

dim ani()
dim aout()
Set rs=Server.CreateObject("ADODB.RecordSet") 
Set prs=Server.CreateObject("ADODB.RecordSet") 
Set srs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from subs where subs_id<>0 " 
if sort_id<>0 and brand_id=0 and subs_id=0 then sql=sql&" and sort_id="& sort_id
if sort_id<>0 and brand_id<>0 and subs_id=0 then sql=sql&" and sort_id="& sort_id &" and  brand_id="&brand_id
if sort_id<>0 and brand_id<>0 and subs_id<>0 then sql=sql&" and subs_id="&subs_id
sql=sql & " and sort_id<>86 order by sort_id , brand_id , code"
rs.open sql,conn,1,1 

%>
<table width="700" border="1" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr bgcolor="#FFFFFF"> 
    <td height="22" colspan="9" align="center" class="header"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr align="center"> 
          <td height="38" colspan="6"><font color="#FF0000"> 
            <%
if pmonth=12 and pday=31 then
response.write("目前")
else
response.write(pmonth&" 月 "&pday&" 日 ")
end if
%>
            </font>的仓库状况</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="80" height="30" bgcolor="#FFFFFF">种类</td>
    <td width="120" bgcolor="#FFFFFF">品牌</td>
    <td width="200" bgcolor="#FFFFFF">型号</td>
    <td width="50" bgcolor="#FFFFFF">原存</td>
    <td width="50" bgcolor="#FFFFFF">现存</td>
    <td width="50" bgcolor="#FFFFFF">欠货</td>
    <td width="50" bgcolor="#FFFFFF">实存</td>
    <td width="50" bgcolor="#FFFFFF">单价</td>
    <td width="50" bgcolor="#FFFFFF">总金额</td>
  </tr>
  <%

redim ain( rs.recordcount - 1 )
redim aout( rs.recordcount - 1 )
i=0

sort_next=0

do while not rs.eof 

if pmonth<>month(now) then psql="select sum(mount) as sum_in from subspass where subspass_type='in' and subs_id=" & rs("subs_id") & " and pmonth<=" & pmonth
if pmonth=month(now) then psql="select sum(mount) as sum_in from subspass where subspass_type='in' and subs_id=" & rs("subs_id") & " and pday<=" & pday
prs.open psql,conn,1,1 
ain(i)=prs("sum_in")
if isnull(prs("sum_in")) then ain(i)=0
prs.close

if pmonth<>month(now) then psql="select sum(mount) as sum_out from subspass where subspass_type='out' and subs_id=" & rs("subs_id") & " and pmonth<=" & pmonth
if pmonth=month(now) then psql="select sum(mount) as sum_out from subspass where subspass_type='out' and subs_id=" & rs("subs_id") & " and pday<=" & pday
prs.open psql,conn,1,1 
aout(i)=prs("sum_out")
if isnull(prs("sum_out")) then aout(i)=0
prs.close

if ( ain(i)- aout(i) + rs("basenum") <> 0 ) or rs("baseown")<>0 then

sort_now=rs("sort_id")
%>
  <tr> 
    <td width="80" height="25" bgcolor="#FFFFFF"> <%
ssql="select * from sort where sort_id="&rs("sort_id") 
srs.open ssql,conn,1,1 
if not(srs.eof or srs.bof) then response.Write(srs("sname")) 
srs.close
%> </td>
    <td width="120" height="25" bgcolor="#FFFFFF"><%
ssql="select * from brand where brand_id="&rs("brand_id") 
srs.open ssql,conn,1,1 
if not(srs.eof or srs.bof) then response.Write(srs("bname"))
srs.close
%></td>
    <td width="200" bgcolor="#FFFFFF"><a href="subs_edit.asp?subs_id=<%=rs("subs_id")%>"><%=rs("code")%><a></td>
    <td width="50" bgcolor="#FFFFFF"><%=rs("basenum")%> <%sbasenum=sbasenum+rs("basenum")%></td>
    <td width="50" bgcolor="#FFFFFF"><%=ain(i)- aout(i) + rs("basenum")%> <%xianchuan=xianchuan+ain(i)- aout(i) + rs("basenum")%></td>
    <td width="50" bgcolor="#FFFFFF"><%=rs("baseown")%> <%sbaseown=sbaseown+rs("baseown")%></td>
    <td width="50" bgcolor="#FFFFFF"><%=ain(i)- aout(i) + rs("basenum") - rs("baseown")%> <%xisu=xisu + ain(i)- aout(i) + rs("basenum") - rs("baseown")%></td>
    <td width="50" bgcolor="#FFFFFF"><%=rs("pricein")%></td>
    <td width="50" bgcolor="#FFFFFF"><%=( ain(i)- aout(i) + rs("basenum") - rs("baseown"))*rs("pricein")%><%xisujin=xisujin + ( ain(i)- aout(i) + rs("basenum") - rs("baseown"))*rs("pricein")%></td>
  </tr>
  <%
end if
rs.movenext
i=i+1                                                                     

if not rs.eof then

if (( ain(i)- aout(i) + rs("basenum") <> 0 ) or rs("baseown")<>0) and ( sbasenum<>0 or xianchuan<>0 or sbaseown<>0 or xisu<>0 )  then 

if sort_now<>rs("sort_id") then

%>
  <tr> 
    <td height="25" colspan="3" bgcolor="#FFFFFF"><font color="#FF0000">总量</font></td>
    <td width="50" bgcolor="#FFFFFF"><font color="#FF0000"><%=sbasenum-zsbasenum%></font></td>
    <td width="50" bgcolor="#FFFFFF"><font color="#FF0000"><%=xianchuan-zxianchuan%></font></td>
    <td width="50" bgcolor="#FFFFFF"><font color="#FF0000"><%=sbaseown-zsbaseown%></font></td>
    <td width="50" bgcolor="#FFFFFF"><font color="#FF0000"><%=xisu-zxisu%></font></td>
    <td width="50" bgcolor="#FFFFFF">&nbsp;</td>
    <td width="50" bgcolor="#FFFFFF"><font color="#FF0000"><%=xisujin-zxisujin%></font></td>
  </tr>
  <%
zsbasenum=sbasenum
zxianchuan=xianchuan
zsbaseown=sbaseown
zxisu=xisu
zxisujin=xisujin
end if
end if
end if
loop
%>
  <tr> 
    <td height="25" colspan="3" bgcolor="#FFFFFF"><font color="#FF0000">总量</font></td>
    <td width="50" bgcolor="#FFFFFF"><font color="#FF0000"><%=sbasenum%></font>&nbsp;</td>
    <td width="50" bgcolor="#FFFFFF"><font color="#FF0000"><%=xianchuan%></font>&nbsp;</td>
    <td width="50" bgcolor="#FFFFFF"><font color="#FF0000"><%=sbaseown%></font>&nbsp;</td>
    <td width="50" bgcolor="#FFFFFF"><font color="#FF0000"><%=xisu%></font>&nbsp;</td>
    <td width="50" bgcolor="#FFFFFF">&nbsp;</td>
    <td width="50" bgcolor="#FFFFFF"><font color="#FF0000"><%=xisujin%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="35" colspan="9">&nbsp;</td>
  </tr>
</table>
<%
rs.close
set rs=nothing
set prs=nothing
set srs=nothing
%>
</body>
</html>
