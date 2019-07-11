<script language="JavaScript">
<!--
function gettarg(targ,selObj,restore){ //v3.0
  eval(targ+".location='user_v.asp?adminlevel="+selObj.options[selObj.selectedIndex].value+"'");
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
adminlevel=request("adminlevel")
if request("option")="editadministratorsucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit administrator is complete. ');</SCRIPT>")
if request("option")="editmanagersucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit manager is complete. ');</SCRIPT>")
if not ( adminlevel=1 or adminlevel=2 or adminlevel=3 or adminlevel=5 ) then adminlevel=1
if adminlevel=1 then 
sql="select * from player where adminlevel=1 order by pid desc"
admin_str="Administrators"
end if
if adminlevel=2 then 
sql="select * from player where adminlevel=2 order by pid desc" 
admin_str="Managers"
end if
if adminlevel=3 then 
sql="select * from player where adminlevel=3 order by pid desc" 
admin_str="Visitors"
end if
if adminlevel=5 then 
sql="select * from player where adminlevel=5 order by pid desc" 
admin_str="Asst.Managers"
end if

Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then 
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' None data ! ');</SCRIPT>")
response.End()
end if
if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=20
rs.AbsolutePage=pagecount
%>
<table width="97%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="22" colspan="8"><select name="adminlevel" id="adminlevel" class="display_dropdown"  onChange="gettarg('self',this,0)">
        <option value=1 selected>Select ...</option>
        <option value=1>Administrator</option>
        <option value=2>Managerial</option>
        <option value=5>Asst.Managerial</option>
        <option value=3>Visitor</option>
      </select></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="22" colspan="8" class="header"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td> <table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
for zi=1 to rs.pagecount
%>
                <td width="50"><a href="user_v.asp?page=<%=zj%>&adminlevel=<%=adminlevel%>">|&nbsp;<%=zj%>&nbsp;|</a></td>
                <%
zj=zj+1
next
%>
              </tr>
            </table></td>
          <td width="20%" align="left"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td> <div align="center"> Record :<font color=red><b><%=rs.recordcount%></b></font> 
                    Page <font color=red><b><%=pagecount%></b></font> of <font color=red><b><%=rs.pagecount%></b></font> 
                  </div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="22" colspan="8" class="header">List 
      of <font color="#FF0000"><%=admin_str%></font></td>
  </tr>
  <tr> 
    <td width="10%" height="22" align="center" bgcolor="#FFFFFF">Userid</td>
    <td width="12%" align="center" bgcolor="#FFFFFF">Password</td>
    <td width="12%" align="center" bgcolor="#FFFFFF">Name</td>
    <td width="14%" align="center" bgcolor="#FFFFFF">Telphone</td>
    <td width="13%" align="center" bgcolor="#FFFFFF">Contact</td>
    <td align="center" bgcolor="#FFFFFF">Email</td>
    <td width="11%" align="center" bgcolor="#FFFFFF">Team</td>
    <td width="15%" align="center" bgcolor="#FFFFFF">OPERATION</td>
  </tr>
  <%
do while not rs.eof 
%>
  <tr> 
    <td width="10%" height="22" align="center" bgcolor="#FFFFFF"><%=rs("uid")%>&nbsp;</td>
    <td width="12%" align="center" bgcolor="#FFFFFF"><%=rs("pwd")%>&nbsp;</td>
    <td width="12%" align="center" bgcolor="#FFFFFF"><%=rs("name")%>&nbsp;</td>
    <td width="14%" align="center" bgcolor="#FFFFFF"><%=rs("tel")%>&nbsp;</td>
    <td width="13%" align="center" bgcolor="#FFFFFF"><%=rs("contact")%>&nbsp;</td>
    <td width="13%" align="center" bgcolor="#FFFFFF"><%=rs("email")%>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF">
		  <%
		  if not isnull(rs("tid")) then
		  if IsNumeric(rs("tid")) then
			tsql="select * from team where tid=" & rs("tid")
			set trs=server.createobject("ADODB.Recordset")
			trs.open tsql,conn,1,1
			if not trs.eof then
		  	response.Write(trs("name"))
			end if
			trs.close
			set trs=nothing
		  end if
		  end if
		  %>&nbsp;
	</td>
    <td align="center" bgcolor="#FFFFFF">
<%
if adminlevel=1 or adminlevel=3 then
%>
	<a href="user_e.asp?id=<%=rs("pid")%>&page=<%=pagecount%>&adminlevel=<%=adminlevel%>">[ EDIT | DEL ]</a> 
<%
elseif adminlevel=2 then
%>
	<a href="manager_e.asp?zurl=user_v.asp&tid=<%=rs("tid")%>&pid=<%=rs("pid")%>&page=<%=pagecount%>&adminlevel=<%=adminlevel%>"> EDIT </a> 
<%
elseif adminlevel=5 then
%>
	<a href="manager_asst_e.asp?zurl=user_v.asp&tid=<%=rs("tid")%>&pid=<%=rs("pid")%>&page=<%=pagecount%>&adminlevel=<%=adminlevel%>"> EDIT </a> 
<%
end if
%>
	 &nbsp;</td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
  <tr bgcolor="#FFFFFF"> 
    <td height="35" colspan="8"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td> <table border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <%
zj=1
for zi=1 to rs.pagecount
%>
                <td width="50"><a href="user_v.asp?page=<%=zj%>&adminlevel=<%=adminlevel%>">|&nbsp;<%=zj%>&nbsp;|</a></td>
                <%
zj=zj+1
next
%>
              </tr>
            </table></td>
          <td width="20%" align="left"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td> <div align="center"> Record :<font color=red><b><%=rs.recordcount%></b></font> Page <font color=red><b><%=pagecount%></b></font> of <font color=red><b><%=rs.pagecount%></b></font> </div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
<%
if adminlevel=2 or adminlevel=5 then
%>
<P><STRONG><FONT face="Verdana, Arial, Helvetica, sans-serif" color=#000066 size=2> 
  <br>
  <br>
  <br>
  <font color="#FF0000">#</font> Because one team must has one manager , <br><br><br>
  the team manager can be edited only.</FONT></STRONG></P><br><br>
<%
end if
%>
</body>
</html>
