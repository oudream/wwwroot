<!--#include file="dbm_adminconn.asp" -->
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
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
Set prs=Server.CreateObject("ADODB.RecordSet")
%>

<table width="98%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <%
sql="select * from dbm_subs where sname<> ''"

if zname<>"" then sql=sql & " and sname like '%"&zname&"%'"
if zsproperties<>"" then sql=sql & " and sproperties like '%"&zsproperties&"%'"
if zstate<>"" then sql=sql & " and sstate like '%"&zstate&"%'"
if zbig<>"" then sql=sql & " and price >= " & zbig
if zsmall<>"" then sql=sql & " and price <= " & zsmall
if zothers<>"" then sql=sql & " and ( subs_id like '%"&zothers&"%' or simg like '%"&zothers&"%' or memo like '%"&zothers&"%')"

sql = sql &" order by sname DESC"

rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' 此用户还没有任何作品 ! ');</SCRIPT>")
%>
  <%
response.End()
end if
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=25
rs.AbsolutePage=pagecount
%>
  <tr align="center"> 
                <td width="60" bgcolor="#FFFFFF">用户</td>
                <td width="100" bgcolor="#FFFFFF">作品名称</td>
                <td width="80" bgcolor="#FFFFFF">尺寸(cm)</td>
                <td width="80" bgcolor="#FFFFFF">市场价(元)</td>
    <td width="51" bgcolor="#FFFFFF">状态</td>
    <td width="148" bgcolor="#FFFFFF">备注</td>
  </tr>
  <%
i=0
do while not rs.eof 
%>
  <tr align="center"> 
                <td bgcolor="#FFFFFF"> 
                  <%
ssql="select * from dbm_guest where guest_id="&rs("guest_id") 
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.open ssql,conn,1,1 
if not(srs.eof or srs.bof) then response.Write(srs("allname")) 
srs.close
	%>
                </td>            
                <td height="25" bgcolor="#FFFFFF"><a href="Show_Me.asp?subs_id=<%=rs("subs_id")%>"><%=rs("sname")%></a> 
                </td>
    <td bgcolor="#FFFFFF"><%=rs("sproperties")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("price")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("sstate")%>&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs("memo")%>&nbsp;</td>
  </tr>
  <%
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
rs.movenext
loop
%>
  <tr bgcolor="#FFFFFF"> 
    <td height="35" colspan="8"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="70%"> <table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
zk=1
for zi=1 to rs.pagecount
if zj<=zk*15 then
%>
                <td width="50" height="30"><a href="ZuoPin.asp?zname=<%=zname%>&zstate=<%=zstate%>&zsproperties=<%=zsproperties%>&zbig=<%=zbig%>&zsmall=<%=zsmall%>&zothers=<%=zothers%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
                <%
else
zk=zk+1
%>
              </tr>
              <tr> 
                <td width="50" height="30"><a href="ZuoPin.asp?zname=<%=zname%>&zstate=<%=zstate%>&zsproperties=<%=zsproperties%>&zbig=<%=zbig%>&zsmall=<%=zsmall%>&zothers=<%=zothers%>&page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
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
</table>
<%
rs.close
set rs=nothing
set prs=nothing
%>
