<SCRIPT LANGUAGE="JavaScript">
<!--

function checkform()
{
	if (zform.zemail.value.length == 0) {
		alert(" Please enter a valid Photo Email . ");
		zform.zemail.focus();
		return false;
		}
	if(! isMail(zform.zemail.value)) {
		alert('Please enter a valid EMAIL address');
		zform.zemail.focus();
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
var letter=" _0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
	if(name.length == 0)
		return false;
	for(i = 0; i < name.length; i++) {
		if(letter.indexOf(name.charAt(i)) == -1)
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
if session("manager_tid")="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select the team you want to edit at first ! ');</SCRIPT>")
response.End()
end if
%>
<%
if trim(request("option"))="del" then
conn.execute "delete * from team where tid=" & session("manager_tid")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del team is complete. ');</SCRIPT>")
response.Redirect("team_v.asp")
response.End()
end if
%>
<%
if trim(request("option"))="edit" then
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
zmpid=request("zmpid")
zapid=request("zapid")
zhcolor=request("zhcolor")
zacolor=request("zacolor")


zdescription=request("zdescription")
if zdescription="" then zdescription=" "
zdescription=replace(zdescription,"'","")
sql="update team set city='"&zcity&"',telfax='"&ztelfax&"',email='"&zemail&"',homepage='"&zhomepage&"',sponsor='"&zsponsor&"',sponsorhomepage='"&zsponsorhomepage&"',createdate='"&zcreatedate&"',photo='"&zphoto&"',hcolor='"&zhcolor&"',acolor='"&zacolor&"',description='"&zdescription&"' where tid=" & session("manager_tid")
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit team is complete. ');</SCRIPT>")
end if
%>

<%
tsql="select * from team where tid=" & session("manager_tid")
set trs=server.createobject("ADODB.Recordset")
trs.open tsql,conn,1,1
if trs.eof or trs.bof then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' You have not join in team ! ');</SCRIPT>")
response.End()
end if
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
      <td height="50" colspan="2"><H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Team Edit</FONT></H2>
        </td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's Name :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ff0000 > 
        <b><%=trs("name")%></b></font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's City :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input  name="zcity" class="form" id="zcity" value="<%=trs("city")%>" size="40" maxlength=100>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's Tel/Fax :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td width="66%">
        <input  name="ztelfax" class="form" id="ztelfax" value="<%=trs("telfax")%>" size="40" maxlength=100>
        </td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's Email :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input  name="zemail" class="form" id="zemail2" value="<%=trs("email")%>" size="40" maxlength=100>
        <font color="#ff0000">*</font> </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's HomePage :</FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input  name="zhomepage" class="form" id="zhomepage" value="<%=trs("homepage")%>" size="40" maxlength=100>
        </font> ( Like &quot;http://www.sunsports.com.sg&quot; )</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's Sponsor :</FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input  name="zsponsor" class="form" id="zsponsor" value="<%=trs("sponsor")%>" size="40" maxlength=100>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Sponsor's HomePage :</FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input  name="zsponsorhomepage" class="form" id="zsponsorhomepage" value="<%=trs("sponsorhomepage")%>" size="40" maxlength=100>
        </font>( Like &quot;http://www.sunsports.com.sg&quot; )</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's Create Date : </FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp;</FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input  name="zcreatedate" class="form" id="zcreatedate" value="<%=trs("createdate")%>" size="40" maxlength=100>
        </font>( Like &quot;14-02-2004&quot; )</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's Photo (Path) : </font><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp;</font></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input  name="zphoto" class="form" id="zphoto" value="<%=trs("photo")%>" size="40" maxlength=100>
        <font color="#ff0000">*</font> </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Manager Nric :</FONT></td>
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

 ======================================================================================================== --></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Asst. Manager Nric :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
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
      </td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="26">&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Home Color :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
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
          <option value=<%=crs("color")%> selected><%=crs("colorname")%></option>
<%
end if
crs.close
set crs=nothing
%>
<%
csql="select * from color"
Set crs=Server.CreateObject("ADODB.RecordSet")
crs.Open csql,conn,1,1
if not ( crs.eof or crs.bof ) then
do while not crs.eof
%>
          <option value=<%=crs("color")%>><%=crs("colorname")%></option>
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
        <font color="#ff0000">*</font></font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Away Color :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
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
          <option value=<%=crs("color")%> selected><%=crs("colorname")%></option>
<%
end if
crs.close
set crs=nothing
%>
          <%
csql="select * from color"
Set crs=Server.CreateObject("ADODB.RecordSet")
crs.Open csql,conn,1,1
if not ( crs.eof or crs.bof ) then
do while not crs.eof
%>
          <option value=<%=crs("color")%>><%=crs("colorname")%></option>
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
        <font color="#ff0000">*</font></font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td valign="top"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >description :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >
        <textarea name="zdescription" cols="50" rows="10" id="zdescription"><%=trs("description")%></textarea>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2">
        <input type="hidden" id="tid" name="tid" value=<%=trs("tid")%>> 
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