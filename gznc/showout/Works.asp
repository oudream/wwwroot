<HTML>
<HEAD>
<TITLE></TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="CSS.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY bgColor=#cccccc leftMargin=0 background=../image/bg_main_005.gif 
topMargin=0 marginheight="0" marginwidth="0">
<!--#include file="dbm_conn.asp" -->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
sql="select * from dbm_subs"
sql = sql &" order by subs_id DESC"
rs.open sql,conn,1,1 
if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' 没有合要求的作品 ! ');</SCRIPT>")
response.End()
end if
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=5
rs.AbsolutePage=pagecount
%>
<table width="660" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="23" rowspan="2">&nbsp;</td>
    <td width="616" height="35" class="P001"> 新筑 &gt; 作品介绍</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td>
<%
do while not rs.eof 
%>
	<table width="100%" border="0" cellspacing="0">
        <tr> 
          <td width="140" height="120" valign="top" bgcolor="#F0F0F0"><img src="<%=rs("sfile")%>" width="140" height="100"></td>
          <td valign="top"><table width="100%" border="0" cellpadding="5" cellspacing="0">
              <tr>
                <td><a href="Works_Show.asp?subs_id=430" class="c"><%=rs("FILETITLE")%></a><br><br>
      <% if len(rs("FILEDESC"))>100 then%>
      <%=left(rs("FILEDESC"),95)%>.. 
      <%else%>
      <%=rs("FILEDESC")%> 
      <%end if%>&nbsp;
				  </td>
              </tr>
            </table></td>
        </tr>
      </table>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
%>
<br>
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
          <td><a href="Works.asp?page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
          <%
else
zk=zk+1
%>
        </tr>
        <tr> 
          <td><a href="Works.asp?page=<%=zj%>">&nbsp;<%=zj%>&nbsp;</a></td>
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
<br>
<%
rs.close
set rs=nothing
%>
      <table width="100%" border="0" cellspacing="0">
        <tr> 
          <td width="140" height="50" valign="top" bgcolor="#F0F0F0">&nbsp;</td>
          <td valign="top">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</BODY></HTML>
