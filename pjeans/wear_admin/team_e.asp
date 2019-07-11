<SCRIPT LANGUAGE="JavaScript">
<!--

function checkform()
{
	if (zform.zname.value.length == 0) {
		alert(" Enter the name and type of team . ");
		zform.zname.focus();
		return false;
		}
	if(! isletter(zform.zname.value)) {
		alert('Please enter a valid Team Name .');
		zform.zname.focus();
		return false;
	}
	if (zform.zphoto.value.length == 0) {
		alert(" Please enter a valid Photo address . ");
		zform.zphoto.focus();
		return false;
		}
	if (zform.zhcolor.value.length == 0) {
		alert(" Please select a valid Home Color . ");
		zform.zhcolor.focus();
		return false;
		}
	if (zform.zacolor.value.length == 0) {
		alert(" Please select a valid Away Color . ");
		zform.zacolor.focus();
		return false;
		}
	return true;
}

function isletter(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charCodeAt(i) > 128)
			return false;
	}
	return true;
}

function isEnglish(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charCodeAt(i) > 128)
			return false;
	}
	return true;
}

function isMail(name)
{
	if(name.length==0)
		return true;
	if(! isEnglish(name))
		return false;
	i = name.indexOf("@");
	j = name.lastIndexOf("@");
	if(i == -1)
		return false;
	if(i != j)
		return false;
	if(i == name.length)
		return false;
	return true;
}

//-->

</SCRIPT>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" -->
<%
if request("option")="editmanagersucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit manager is complete. ');</SCRIPT>")
if request("option")="editmanagerasstsucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit Asst.manager is complete. ');</SCRIPT>")
if request("tid")="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select the team you want to edit at first ! ');</SCRIPT>")
response.End()
end if
%>
<%
if trim(request("option"))="edit" then
zname=trim(request("zname"))
if zname<>"" then

zhcolor=split(trim(request("zhcolor")),"zooz",-1,1)
if ubound(zhcolor)<>1 then
response.write("<SCRIPT LANGUAGE=JavaScript> alert(' The Home Color is error ! ');</SCRIPT>")
response.End()
end if
zhcolora=zhcolor(0)
zhcolorb=zhcolor(1)

zacolor=split(trim(request("zacolor")),"zooz",-1,1)
if ubound(zacolor)<>1 then
response.write("<SCRIPT LANGUAGE=JavaScript> alert(' The Away Color is error ! ');</SCRIPT>")
response.End()
end if
zacolora=zacolor(0)
zacolorb=zacolor(1)

tsql="select * from team where name='"&zname&"' and tid<>" & request("tid")
set trs=server.createobject("ADODB.Recordset")
trs.open tsql,conn,1,1
if not ( trs.eof or trs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' The Team name have existed , please enter other Team name . ');</SCRIPT>")
trs.close
set trs=nothing
else 
trs.close
set trs=nothing

zcity=request("zcity")
if zcity="" then zcity=" "
ztelfax=request("ztelfax")
if ztelfax="" then ztelfax=" "
zemail=request("zemail")
if zemail="" then zemail=" "
zhomepage=request("zhomepage")
if zhomepage="" then zhomepage=" "
zsponsor=request("zsponsor")
if zsponsor="" then zsponsor=" "
zsponsorhomepage=request("zsponsorhomepage")
if zsponsorhomepage="" then zsponsorhomepage=" "
zcreatedate=request("zcreatedate")
if zcreatedate="" then zcreatedate=" "

zphoto=request("zphoto")

zdescription=request("zdescription")
if zdescription="" then zdescription=" "
zdescription=replace(zdescription,"'","¡¯")
conn.Execute("update team set name='"&zname&"',city='"&zcity&"',telfax='"&ztelfax&"',email='"&zemail&"',homepage='"&zhomepage&"',sponsor='"&zsponsor&"',sponsorhomepage='"&zsponsorhomepage&"',createdate='"&zcreatedate&"',photo='"&zphoto&"',hcolor='"&zhcolora&"',hcolorname='"&zhcolorb&"',acolor='"&zacolora&"',acolorname='"&zacolorb&"',description='"&zdescription&"' where tid=" & request("tid"))
conn.Execute("update card set tname='"&zname&"' where tid=" & request("tid"))
conn.Execute("update scorer set tname='"&zname&"' where tid=" & request("tid"))
conn.Execute("update standing set tname='"&zname&"' where tid=" & request("tid"))
conn.Execute("update match set htname='"&zname&"' where htid=" & request("tid"))
conn.Execute("update match set atname='"&zname&"' where atid=" & request("tid"))
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit team is complete. ');</SCRIPT>")
end if
end if
end if 
%>

<%
tsql="select * from team where tid=" & request("tid")
set trs=server.createobject("ADODB.Recordset")
trs.open tsql,conn,1,1
%> 

<form id="zform" action="team_e.asp" onSubmit="return checkform();" method="post">
  <table width="100%" border="0" cellspacing="2" cellpadding="2">
    <tr> 
      <td width="1%">&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td width="4%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td height="50" colspan="2" align="center"><H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066> Edit Team</FONT></H2></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="38" colspan="4"><strong><font color="#FF0000">Step 1</font></strong></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>Manager Nric :</FONT></td>
      <td width="66%"> 
        <!-- ========================================================================================================		  

													Manager Nric ============star

 ======================================================================================================== -->
        <%
psql="select * from player where pid=" & trs("mpid")
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if not ( prs.bof or prs.eof ) then 
response.Write(prs("name"))
end if
prs.close
set prs=nothing
%> 
        <!-- ========================================================================================================		  

													Manager Nric ============stop

 ======================================================================================================== -->
        <a href="manager_e.asp?zurl=team_e.asp&tid=<%=trs("tid")%>&pid=<%=trs("mpid")%>">( 
        Click Here to Edit the Manager ) </a></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>Asst. Manager Nric :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff>&nbsp; </FONT></td>
      <td width="66%"> 
        <!-- ========================================================================================================		  

													Manager Nric ============star

 ======================================================================================================== -->
        <%
psql="select * from player where pid=" & trs("apid")
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if not ( prs.bof or prs.eof ) then 
response.Write(prs("name"))
end if
prs.close
set prs=nothing
%> 
        <!-- ========================================================================================================		  

													Manager Nric ============stop

 ======================================================================================================== -->
        <a href="manager_asst_e.asp?zurl=team_e.asp&tid=<%=trs("tid")%>&pid=<%=trs("apid")%>">( 
        Click Here to Edit the Asst. Manager ) </a></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="38" colspan="4"><strong><font color="#FF0000">Step 2</font></strong></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>Team's Name :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff>&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff> 
        <input id="zname"  name="zname" maxlength=1000 value="<%=trs("name")%>" class="form">
        <font color="#000000">*</font> </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>Team's City :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff>&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff> 
        <input id="zcity"  name="zcity" maxlength=1000 value="<%=trs("city")%>" class="form">
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>Team's Tel/Fax :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff>&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff> 
        <input id="ztelfax"  name="ztelfax" maxlength=1000 value="<%=trs("telfax")%>" class="form">
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>Team's Email :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff>&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff> 
        <input id="zemail"  name="zemail" maxlength=1000 value="<%=trs("email")%>" class="form">
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>Team's HomePage :</FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff> 
        <input id="zhomepage"  name="zhomepage" maxlength=1000 value="<%=trs("homepage")%>" class="form">
        </font> ( Like &quot;http://www.sunsports.com.sg&quot; )</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>Team's Sponsor :</FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff> 
        <input id="zsponsor"  name="zsponsor" maxlength=1000 value="<%=trs("sponsor")%>" class="form">
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>Sponsor's HomePage :</FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff> 
        <input id="zsponsorhomepage"  name="zsponsorhomepage" maxlength=1000 value="<%=trs("sponsorhomepage")%>" class="form">
        </font>( Like &quot;http://www.suosports.com.sg&quot; )</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>Team's Create Date : </FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff>&nbsp;</FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff> 
        <input id="zcreatedate"  name="zcreatedate" maxlength=1000 value="<%=trs("createdate")%>" class="form">
        </font>( Like &quot;14-02-2004&quot; )</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>Team's Photo (Path) : </font><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff>&nbsp;</font></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff> 
        <input id="zphoto"  name="zphoto" maxlength=1000 value="<%=trs("photo")%>" class="form">
        <font color="#000000">*</font> </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="26">&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>Home Color :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff>&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff> 
        <!-- ========================================================================================================		  

													Manager Nric ============star

 ======================================================================================================== -->
        <select name="zhcolor" id="zhcolor" class="display_dropdown">
          <%
csql="select * from color where color='" & trs("hcolor") & "'"
Set crs=Server.CreateObject("ADODB.RecordSet")
crs.Open csql,conn,1,1
if not ( crs.eof or crs.bof ) then
%>
          <option value="<%=crs("color")&"zooz"&crs("colorname")%>" selected><%=crs("colorname")%></option>
          <%
end if
crs.close
set crs=nothing
%>
          <%
csql="select * from color where color<>'" & trs("hcolor") & "'"
Set crs=Server.CreateObject("ADODB.RecordSet")
crs.Open csql,conn,1,1
if not ( crs.eof or crs.bof ) then
do while not crs.eof
%>
          <option value="<%=crs("color")&"zooz"&crs("colorname")%>"><%=crs("colorname")%></option>
          <%
crs.movenext
loop
end if
crs.close
set crs=nothing
%>
        </select>
        <!-- ========================================================================================================		  

													Manager Nric ============stop

 ======================================================================================================== -->
        <font color="#000000">*</font></font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>Away Color :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff>&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff> 
        <!-- ========================================================================================================		  

													Manager Nric ============star

 ======================================================================================================== -->
        <select name="zacolor" id="zacolor" class="display_dropdown">
          <%
csql="select * from color where color='" & trs("acolor") & "'"
Set crs=Server.CreateObject("ADODB.RecordSet")
crs.Open csql,conn,1,1
if not ( crs.eof or crs.bof ) then
%>
          <option value="<%=crs("color")&"zooz"&crs("colorname")%>" selected><%=crs("colorname")%></option>
          <%
end if
crs.close
set crs=nothing
%>
          <%
csql="select * from color where color<>'" & trs("acolor") & "'"
Set crs=Server.CreateObject("ADODB.RecordSet")
crs.Open csql,conn,1,1
if not ( crs.eof or crs.bof ) then
do while not crs.eof
%>
          <option value="<%=crs("color")&"zooz"&crs("colorname")%>"><%=crs("colorname")%></option>
          <%
crs.movenext
loop
end if
crs.close
set crs=nothing
%>
        </select>
        <!-- ========================================================================================================		  

													Manager Nric ============stop

 ======================================================================================================== -->
        <font color="#000000">*</font></font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>description :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff>&nbsp; </FONT></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff> 
        <textarea name="zdescription" cols="50" rows="10" id="zdescription"><%=trs("description")%></textarea>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2"> <input type="hidden" id="tid" name="tid" value=<%=trs("tid")%>> 
        <input type="submit" id="option" name="option" value="edit"> 
        <input type="button" name="goback" value="go back" onClick="history.go( -1 );return true;"> 
      </td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td height="50" colspan="2">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<%
trs.close
set trs=nothing
%>