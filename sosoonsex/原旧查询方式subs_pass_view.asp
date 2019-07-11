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
if trim(request("option"))="del" then
conn.execute "delete * from subspass where subspass_id=" & request("subspass_id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('删除成功');</SCRIPT>")
end if
%>
<form name="form1" method="post" action="subs_pass_view.asp">
        <table width="98%" border="1" align="center" cellpadding="0" cellspacing="0">
          <tr>
            
      <td width="100" height="31">流通类型</td>
            <td width="100">客户</td>
            <td width="150">商品</td>
            <td width="150">日期</td>
            <td width="100">本公司经手人</td>
            <td width="50">&nbsp;</td>
          </tr>
          <tr> 
            <td width="100"> <select name="subspass_type" id="subspass_type" class="display_dropdown">
                <option value="" selected>请选择……</option>
                <option value="in">进货</option>
                <option value="out">出货</option>
              </select></td>
            
      <td width="100"> <select name="gid" id="gid" class="display_dropdown">
          <option value=0 selected> 请选择……</option>
          <%
gsql="select * from guest"
Set grs=Server.CreateObject("ADODB.RecordSet")
grs.Open gsql,conn,1,1
if not ( grs.eof or grs.bof ) then
do while not grs.eof
%>
          <option value=<%=grs("gid")%>><%=grs("simname")%></option>
          <%
grs.movenext
loop
end if
grs.close
set grs=nothing
%>
        </select></td>
            <td width="150"> <select name="subs_id" id="subs_id" class="display_dropdown">
                <option value=0 selected> 请选择……</option>
                <%
ssql="select * from subs "
Set srs=Server.CreateObject("ADODB.RecordSet")
srs.Open ssql,conn,1,1
do while not srs.eof
bsql="select bname from brand where brand_id="&srs("brand_id")
Set brs=Server.CreateObject("ADODB.RecordSet")
brs.Open bsql,conn,1,1
if not ( brs.eof or brs.bof ) then
%>
                <option value=<%=srs("subs_id")%>><%=brs("bname")&" "&srs("code")%></option>
                <%
end if
brs.close
srs.movenext
loop
srs.close
set srs=nothing
%>
              </select></td>
            <td width="150"> <input name="pmonth" id="pmonth" type="text" size="3" maxlength="2">
              月 
              <input name="pday" id="pday" type="text" size="3" maxlength="2">
              日 </td>
            <td width="100"><select name="operationer_id" id="operationer_id" class="display_dropdown">
                <option value=0 selected> 请选择……</option>
                <%
psql="select * from user"
Set prs=Server.CreateObject("ADODB.RecordSet")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
                <option value=<%=prs("id")%>><%=prs("uname")%></option>
                <%
prs.movenext
loop
end if
prs.close
set prs=nothing
%>
              </select></td>
            <td width="50"> <input type="submit" name="Submit" value="确定"></td>
          </tr>
        </table>
      </form> 

<%
subspass_type=request("subspass_type")
gid=request("gid")
if gid="" then gid=0
subs_id=request("subs_id")
if subs_id="" then subs_id=0
pmonth=request("pmonth")
pday=request("pday")
operationer_id=request("operationer_id")
if operationer_id="" then operationer_id=0

sql="select * from subspass where subspass_id<>0" 
if subspass_type<>"" then sql=sql&" and subspass_type='"&subspass_type&"'"
if gid<>0 then sql=sql&" and gid="&gid
if subs_id<>0 then sql=sql&" and subs_id="&subs_id
if pmonth<>"" then sql=sql&" and pmonth="&pmonth
if pday<>"" then sql=sql&" and pday="&pday
if operationer_id<>0 then sql=sql&" and operationer_id="&operationer_id
sql=sql&" order by pday desc,pmonth desc,gid desc,subs_id desc,subspass_type desc,operationer_id desc"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('没有所要的商品');</SCRIPT>")
else
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=20
rs.AbsolutePage=pagecount
%>
<table width="98%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr bgcolor="#FFFFFF"> 
    <td height="22" colspan="9" align="center" class="header"><table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr align="center"> 
          <td height="38" colspan="6">所要查询的<font color="#FF0000">结果</font></td>
        </tr>
      </table></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="22" colspan="9" align="center" class="header"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="75%"> <table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
for zi=1 to rs.pagecount
%>
                <td width="50"><a href="subs_pass_view.asp?page=<%=zj%>&subspass_type=<%=subspass_type%>&gid=<%=gid%>&subs_id=<%=subs_id%>&pmonth=<%=pmonth%>&pday=<%=pday%>&operationer_id=<%=operationer_id%>">&nbsp;<%=zj%>&nbsp;</a></td>
                <%
zj=zj+1
next
%>
              </tr>
            </table></td>
          <td width="24%" align="left"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td> <div align="center"> Record :<font color=red><b><%=rs.recordcount%></b></font> Page <font color=red><b><%=pagecount%></b></font> of <font color=red><b><%=rs.pagecount%></b></font> </div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="40" height="30" bgcolor="#FFFFFF">流通类型</td>
    <td width="150" bgcolor="#FFFFFF">商品</td>
    <td width="40" bgcolor="#FFFFFF">单价</td>
    <td width="40" bgcolor="#FFFFFF">数量</td>
    <td width="80" bgcolor="#FFFFFF">日期</td>
    <td width="70" bgcolor="#FFFFFF">经手人</td>
    <td width="80" bgcolor="#FFFFFF">客户</td>
    <td width="90" bgcolor="#FFFFFF">备注</td>
    <td width="80" bgcolor="#FFFFFF">Operation</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr> 
    <td width="40" height="25" bgcolor="#FFFFFF"><%
   	Select Case rs("subspass_type")
           Case "in"    response.Write("进货")
           Case "out"   response.Write("出货")
      End Select
%>&nbsp;</td>
    <td width="150" height="25" bgcolor="#FFFFFF"> <%
ssql="select * from subs where subs_id=" & rs("subs_id")
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.open ssql,conn,1,1 
if not ( srs.bof or srs.eof ) then 
bsql="select * from brand where brand_id=" & srs("brand_id")
Set brs=Server.CreateObject("ADODB.RecordSet") 
brs.open bsql,conn,1,1 
if not ( brs.bof or brs.eof ) then 
response.Write(brs("bname")&"  "&srs("code"))
brs.close
end if
srs.close
end if
%> 
      &nbsp;</td>
    <td width="40" bgcolor="#FFFFFF"><%=rs("price")%>&nbsp;</td>
    <td width="40" bgcolor="#FFFFFF"><%=rs("mount")%>&nbsp;</td>
    <td width="80" bgcolor="#FFFFFF"><%=rs("pmonth")%> 月 <%=rs("pday")%>日 </td>
    <td width="70" bgcolor="#FFFFFF">
<%
if rs("operationer_id")<>0 then
usql="select * from user where id=" & rs("operationer_id")
Set urs=Server.CreateObject("ADODB.RecordSet") 
urs.open usql,conn,1,1 
if not ( urs.bof or urs.eof ) then 
response.Write(urs("uname"))
end if
urs.close
end if
%>
    </td>
    <td width="80" bgcolor="#FFFFFF">
<%
Set zrs=Server.CreateObject("ADODB.RecordSet")
zsql="select simname from guest where gid="&rs("gid")
zrs.Open zsql,conn,1,1
if not ( zrs.bof or zrs.eof ) then
response.Write(zrs("simname"))
end if
zrs.close
set zrs=nothing
%>
	&nbsp;</td>
    <td width="90" bgcolor="#FFFFFF"><%=rs("memo")%>&nbsp;</td>
    <td width="80" bgcolor="#FFFFFF"><a href="subs_pass_edit.asp?subspass_id=<%=rs("subspass_id")%>">修改</a>&nbsp;&nbsp;&nbsp;<a href="subs_pass_view.asp?option=del&subspass_id=<%=rs("subspass_id")%>" onClick="return confirm('确定删除此项吗？')">删除</a></td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
  <tr bgcolor="#FFFFFF"> 
    <td height="35" colspan="9">&nbsp;</td>
  </tr>
</table>
<%
end if
rs.close
set rs=nothing
%>
</body>
</html>
