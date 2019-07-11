<LINK href="css.css" type=text/css rel=stylesheet>
<SCRIPT LANGUAGE="JavaScript">
<!--

function checkform()
{
	if (zform.zname.value.length == 0) {
		alert(" Enter the name and type of team . ");
		zform.zname.focus();
		return false;
		}
	if(! isEnglish(zform.zname.value)) {
		alert('Enter the name and type of team .');
		zform.zname.focus();
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
	if (zform.zmpid.value.length == 0) {
		alert(" Please select a valid Manager Nric . ");
		zform.zmpid.focus();
		return false;
		}
	if (zform.zapid.value.length == 0) {
		alert(" Please select a valid Asst. Manager Nric . ");
		zform.zapid.focus();
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
<body text="#000000" leftmargin="8" topmargin="8" marginwidth="0" marginheight="0">

<!--#include file="adminconn.asp" -->
<%
if request("option")="creammangersucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create manager is complete. ');</SCRIPT>")
if request("option")="creammangerasstsucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create Asst.manager is complete. ');</SCRIPT>")
if request("option")="add" then

zname=trim(request("zname"))
zphoto=request("zphoto")
zmpid=request("zmpid")
zapid=request("zapid")
cid=request("cid")

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

tsql="select * from team where name='"&zname&"'"
set trs=server.createobject("ADODB.Recordset")
trs.open tsql,conn,1,1
if not ( trs.eof or trs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' The Team name have existed , please enter other Team name . ');</SCRIPT>")
trs.close
set trs=nothing
else 
trs.close
set trs=nothing

psql="select * from player where pid=" & zmpid
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1
if not isnull(prs("tid")) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' The Manager has already managed other teams,please create manager again or choose again . ');</SCRIPT>")
prs.close
set prs=nothing
else
prs.close
set prs=nothing

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
zdescription=request("zdescription")
if zdescription="" then zdescription=" "
zdescription=replace(zdescription,"'","''")

conn.Execute("insert into team (name,city,telfax,email,homepage,sponsor,sponsorhomepage,createdate,photo,mpid,apid,hcolor,hcolorname,acolor,acolorname,description,inserttime) values ('"&zname&"','"&zcity&"','"&ztelfax&"','"&zemail&"','"&zhomepage&"','"&zsponsor&"','"&zsponsorhomepage&"','"&zcreatedate&"','"&zphoto&"','"&zmpid&"','"&zapid&"','"&zhcolora&"','"&zhcolorb&"','"&zacolora&"','"&zacolorb&"','"&zdescription&"','"&now&"')")

sql="select * from team order by tid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
conn.Execute("update player set tid="&rs("tid")&" where pid=" & zmpid)
conn.Execute("update player set tid="&rs("tid")&" where pid=" & zapid)
if cid<>"" then conn.Execute("insert into ct ( cid ,tid ) values ( "& request("cid")&","&rs("tid") &")")
rs.close
set rs=nothing

response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create team is complete. ');</SCRIPT>")
end if
end if
end if
%>

<form id="zform" action="team_c.asp" onSubmit="return checkform();" method="post">
  <table width="100%" border="0" cellspacing="2" cellpadding="2">
    <tr> 
      <td width="1%">&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td width="4%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td height="50" colspan="2"><H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Create Team Wizard </FONT></H2></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="38" colspan="4"><strong><font color="#FF0000">Step 1</font></strong></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="25%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Manager:</FONT></td>
      <td width="70%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <!-- ========================================================================================================		  

													Manager Nric ============star

 ======================================================================================================== -->
        <select name="zmpid" id="zmpid" class="display_dropdown">
          <%
psql="select * from player where adminlevel=2 order by pid"
Set prs=Server.CreateObject("ADODB.RecordSet")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
          <option value=<%=prs("pid")%> selected><%=prs("name")%> -- 
		  <%
		  if not isnull(prs("tid")) then
		  if IsNumeric(prs("tid")) then
			tsql="select * from team where tid=" & prs("tid")
			set trs=server.createobject("ADODB.Recordset")
			trs.open tsql,conn,1,1
			if not trs.eof then
		  	response.Write(trs("name"))
			end if
			trs.close
			set trs=nothing
		  end if
		  end if
		  %>
		  </option>
          <%
prs.movenext
loop
end if
prs.close
set prs=nothing
%>
        </select>
        <!-- ========================================================================================================		  

													Manager Nric ============stop

 ======================================================================================================== -->
        <font color="#FF0000">*</font> </font><a href="manager_c.asp">(If No Tihs 
        Player , <font color="#FF0000">Click Here</font> to Creat The New Manager 
        ) </a></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="25%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Asst. Manager:</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td width="70%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <!-- ========================================================================================================		  

													Manager Nric ============star

 ======================================================================================================== -->
        <select name="zapid" id="zapid" class="display_dropdown">
          <%
psql="select * from player where ptype='a' order by pid"
Set prs=Server.CreateObject("ADODB.RecordSet")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
          <option value=<%=prs("pid")%> selected><%=prs("name")%> -- 
		  <%
		  if not isnull(prs("tid")) then
		  if IsNumeric(prs("tid")) then
			tsql="select * from team where tid=" & prs("tid")
			set trs=server.createobject("ADODB.Recordset")
			trs.open tsql,conn,1,1
			if not trs.eof then
		  	response.Write(trs("name"))
			end if
			trs.close
			set trs=nothing
		  end if
		  end if
		  %>
		  </option>
          <%
prs.movenext
loop
end if
prs.close
set prs=nothing
%>
        </select>
        <!-- ========================================================================================================		  

													Manager Nric ============stop

 ======================================================================================================== -->
        <font color="#FF0000">*</font> </font><a href="manager_asst_c.asp">( If 
        No This Player , <font color="#FF0000">Click Here</font> to Creat The 
        New Asst.Manager ) </a></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="38" colspan="4"><strong><font color="#FF0000">Step 2</font></strong></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="25%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's Name :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td width="70%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input id="zname"  name="zname" maxlength=100 0>
        <font color="#FF0000">*</font> </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="25%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's City :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td width="70%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input id="zcity"  name="zcity" maxlength=100 0>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="25%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's Tel/Fax :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td width="70%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input id="ztelfax"  name="ztelfax" maxlength=100 0>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="25%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's Email :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td width="70%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input id="zemail"  name="zemail" maxlength=100 0>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="25%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's HomePage :</FONT></td>
      <td width="70%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input id="zhomepage"  name="zhomepage" maxlength=100 0>
        </font> ( Like &quot;http://www.dnyes.com&quot; )</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="25%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's Sponsor :</FONT></td>
      <td width="70%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input id="zsponsor"  name="zsponsor" maxlength=100 0>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="25%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Sponsor's HomePage :</FONT></td>
      <td width="70%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input id="zsponsorhomepage"  name="zsponsorhomepage" maxlength=100 0>
        </font>( Like &quot;http://www.dnyes.com&quot; )</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="25%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's Create Date : </FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp;</FONT></td>
      <td width="70%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input id="zcreatedate"  name="zcreatedate" maxlength=100 0>
        </font>( Like &quot;14-02-2004&quot; )</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="25%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Team's Photo (Path) : </font><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp;</font></td>
      <td width="70%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input id="zphoto"  name="zphoto" maxlength=100 0 value="team_images/">
        <font color="#FF0000">*</font> </font><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >( Photo in &quot;team_images/ &quot; )</FONT></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="27">&nbsp;</td>
      <td width="25%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Home Color :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td width="70%">
        <!-- ========================================================================================================		  

													Manager Nric ============star

 ======================================================================================================== -->
        <select name="zhcolor" id="zhcolor" class="display_dropdown">
          <option value="" selected> SELECT ...</option>
          <%
csql="select * from color order by colorname"
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
        <font color="#FF0000">*</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="25%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Away Color :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td width="70%">
        <!-- ========================================================================================================		  

													Manager Nric ============star

 ======================================================================================================== -->
        <select name="zacolor" id="zacolor" class="display_dropdown">
          <option selected value="">SELECT ...</option>
          <%
csql="select * from color order by colorname"
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
        <font color="#FF0000">*</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >description :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <textarea name="zdescription" cols="50" rows="10" id="zdescription"></textarea>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2"><strong><font color="#FF0000">Step 3</font></strong></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Choose a profile :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >
        <select name="cid" class="display_dropdown" id="cid" >
                            <option selected value="">SELECT ... </option>
                            <%
sql="select * from tournament order by type desc, cid" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
                            <option value=<%=rs("cid")%>><%=rs("name")%> </option>
                            <%
rs.movenext
loop
rs.close
set rs=nothing
%>
                          </select>
				   </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2"> <INPUT type=submit id="option" name="option" value="add"> 
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
<br>
<br>
</body>
</html>
