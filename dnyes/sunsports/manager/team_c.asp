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
		alert(' Sorry ! the name of team makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
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

<!--#include file="adminconn.asp" -->
<%
zname=trim(request("zname"))
if zname<>"" then
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
zdescription=replace(zdescription,"'","¡¯")
sql="insert into team (name,city,telfax,email,homepage,sponsor,sponsorhomepage,createdate,photo,mpid,apid,hcolor,acolor,description,inserttime) values ('"&zname&"','"&zcity&"','"&ztelfax&"','"&zemail&"','"&zhomepage&"','"&zsponsor&"','"&zsponsorhomepage&"','"&zcreatedate&"','"&zphoto&"','"&zmpid&"','"&zapid&"','"&zhcolor&"','"&zacolor&"','"&zdescription&"','"&now&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create team is complete. ');</SCRIPT>")
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
                  color=#000066>Team Wizard </FONT></H2>
        </td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Team's Name :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="zname"  name="zname" maxlength=100 size=30>
        <font color="#000000">*</font> </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Team's City :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="zcity"  name="zcity" maxlength=100 size=30>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Team's Tel/Fax :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="ztelfax"  name="ztelfax" maxlength=100 size=30>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Team's Email :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="zemail"  name="zemail" maxlength=100 size=30>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Team's HomePage :</FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="zhomepage"  name="zhomepage" maxlength=100 size=30>
        </font> ( Like &quot;http://www.dnyes.com&quot; )</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Team's Sponsor :</FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="zsponsor"  name="zsponsor" maxlength=100 size=30>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Sponsor's HomePage :</FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="zsponsorhomepage"  name="zsponsorhomepage" maxlength=100 size=30>
        </font>( Like &quot;http://www.dnyes.com&quot; )</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Team's Create Date : </FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp;</FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="zcreatedate"  name="zcreatedate" maxlength=100 size=30>
        </font>( Like &quot;14-02-2004&quot; )</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Team's Photo (Path) : </font><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp;</font></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="zphoto"  name="zphoto" maxlength=100 size=30>
        <font color="#000000">*</font> </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Manager Nric :</FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <!-- ========================================================================================================		  

													Manager Nric ============star

 ======================================================================================================== -->
        <select name="zmpid" id="zmpid" class="display_dropdown">
          <option selected value="">SELECT ...</option>
          <%
psql="select * from player order by name"
Set prs=Server.CreateObject("ADODB.RecordSet")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
          <option value=<%=prs("pid")%>><%=prs("name")%></option>
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
        <font color="#000000">*</font> </font><a href="player_c.asp?zurl=team_c.asp">( Click Here 
        to register New Player ) </a></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Asst. Manager Nric :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <!-- ========================================================================================================		  

													Manager Nric ============star

 ======================================================================================================== -->
        <select name="zapid" id="zapid" class="display_dropdown">
          <option selected value="">SELECT ...</option>
          <%
psql="select * from player order by name"
Set prs=Server.CreateObject("ADODB.RecordSet")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
          <option value=<%=prs("pid")%>><%=prs("name")%></option>
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
        <font color="#000000">*</font> </font><a href="player_c.asp?zurl=team_c.asp">( Click Here 
        to register New Player ) </a></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="27">&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Home Color :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <!-- ========================================================================================================		  

													Manager Nric ============star

 ======================================================================================================== -->
        <select name="zhcolor" id="zhcolor" class="display_dropdown">
          <option selected value="">SELECT ...</option>
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
        <font color="#000000">*</font></font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="29%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Away Color :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <!-- ========================================================================================================		  

													Manager Nric ============star

 ======================================================================================================== -->
        <select name="zacolor" id="zacolor" class="display_dropdown">
          <option selected value="">SELECT ...</option>
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
        <font color="#000000">*</font></font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>description :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>
        <textarea name="zdescription" cols="50" rows="10" id="zdescription"></textarea>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2"><input type=submit value=Submit name=submit></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066 size=2><STRONG>* Mandatory Fields</STRONG></FONT> 
      </td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
