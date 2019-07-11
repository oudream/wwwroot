<script language="JavaScript">

function checkform()
{
	if (form_administrator.name.value.length == 0) {
		alert("Please enter a valid Name. ");
		form_administrator.uname.focus();
		return false;
		}
	if(! isLetter(form_administrator.name.value)) {
		alert("Please enter a valid Name. ");
		form_administrator.name.focus();
		return false;
		}
	if(! isMail(form_administrator.email.value)) {
		alert('Please enter a valid EMAIL address');
		form_administrator.email.focus();
		return false;
		}
	if (form_administrator.subject.value.length == 0) {
		alert("Please enter a valid title. ");
		form_administrator.subject.focus();
		return false;
		}
	if(! isLetter(form_administrator.subject.value)) {
		alert("Please enter a valid title. ");
		form_administrator.subject.focus();
		return false;
		}
	if (form_administrator.mailtext.value.length == 0) {
		alert("Please enter a valid content. ");
		form_administrator.mailtext.focus();
		return false;
		}
	if(! isLetter(form_administrator.mailtext.value)) {
		alert("Please enter a valid content. ");
		form_administrator.mailtext.focus();
		return false;
		}
	return true;
}

function isLetter(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charCodeAt(i) > 128)
			return false;
	}
	return true;
}

function isNumber(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charAt(i) < "0" || name.charAt(i) > "9")
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

</script>

<!--#include file="Top_sun.asp"-->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181" valign="top"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td> 
            <!-- ========================================================================================================		  

													cup table ============star

 ======================================================================================================== -->
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="League.asp"><b><font color="#1F4A71">LEAGUE MATCHES</font></b></a><font color="#1F4A71">+ 
                  +</font></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Cup.asp"><b><font color="#1F4A71">CUP MATCHES </font></b></a><font color="#1F4A71">&nbsp;+ 
                  +</font></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Friendlies.asp"><b><font color="#1F4A71">FRIENDLY 
                  MATCHES </font></b></a></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="TeamProfile.asp"><b><font color="#1F4A71">TEAMPORFILE</font></b></a> 
                  <font color="#1F4A71">&nbsp;+ +</font> </td>
              </tr>
            </table>
            <!-- ========================================================================================================		  

													cup table ============stop

 ======================================================================================================== -->
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table></td>
    <td width="20"> </td>
    <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="27">&nbsp;</td>
        </tr>
        <tr> 
          <td height="50" valign="top" class="header_b">Contact Us</td>
        </tr>
        <tr> 
          <td>Please send us an <a href="mailto:sunsport@singnet.com.sg" class="b">email</a> 
            if you have any queries with regards to participation in the various 
            soccer leagues. Alternatively, please use the form below to submit 
            any queries or comments you may have.<br></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table>
      <%
if request("option")="add" then
uname=trim(request("name"))
uname=replace(uname,"'","''")
email=trim(request("email"))
subject=trim(request("subject"))
subject=replace(subject,"'","''")
mailtext=trim(request("mailtext"))
mailtext=replace(mailtext,"'","''")
sql="insert into mail (name,email,subject,mailtext,inserttime) values ('"&uname&"','"&email&"','"&subject&"','"&mailtext&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Your Email is send , we will reply you as quickly as possible . ');</SCRIPT>")
end if
%>
<FORM action="contact.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
                        
        <TABLE width="100%" border=0 align="center" cellPadding=0 
                        cellSpacing=0 bgColor=#CCCCCC>
          <TBODY>
            <TR bgColor=#ffffff> 
              <TD bgColor=#ffffff width="75%"> <INPUT class=form id="name"
                                onblur="if(this.value=='')this.value='Name';" 
                                style="FONT-SIZE: 12px; WIDTH: 360px; background:white;border:1px solid #6184A3" onMouseOver="this.style.border='1px solid red'" onMouseOut="this.style.border='1px solid #6184A3'" 
                                onfocus="if(this.value=='Name')this.value='';" 
                                value=Name name=name
								
								> </TD>
            </TR>
            <TR bgColor=#ffffff> 
              <TD width="75%" height="27" bgColor=#ffffff> <INPUT class=form id="email"
                                onblur="if(this.value=='')this.value='Email';" 
                                style="FONT-SIZE: 12px; WIDTH: 360px; background:white;border:1px solid #6184A3" onMouseOver="this.style.border='1px solid red'" onMouseOut="this.style.border='1px solid #6184A3'" 
                                onfocus="if(this.value=='Email')this.value='';" 
                                value=Email name=email> </TD>
            </TR>
            <TR bgColor=#ffffff> 
              <TD bgColor=#ffffff width="75%"> <INPUT class=form id="subject"
                                onblur="if(this.value=='')this.value='Subject';" 
                                style="FONT-SIZE: 12px; WIDTH: 360px; background:white;border:1px solid #6184A3" onMouseOver="this.style.border='1px solid red'" onMouseOut="this.style.border='1px solid #6184A3'" 
                                onfocus="if(this.value=='Subject')this.value='';" 
                                value=Subject name=subject> </TD>
            </TR>
            <TR bgColor=#ffffff> 
              <TD bgColor=#ffffff width="75%"> <TEXTAREA id="mailtext" class=form onblur="if(this.value=='')this.value='Mailtext';"                                 style="FONT-SIZE: 12px; WIDTH: 360px; background:white;border:1px solid #6184A3" onMouseOver="this.style.border='1px solid red'" onMouseOut="this.style.border='1px solid #6184A3'" 
 onfocus="if(this.value=='Mailtext')this.value='';" name=mailtext rows=10 wrap=VIRTUAL value="Mailtext">Mailtext</TEXTAREA> 
              </TD>
            </TR>
            <TR bgColor=#ffffff> 
              <TD bgColor=#ffffff> <input type=submit id="options" name="options" value="          [ Submit Here ]          " class="button"> 
                <input type="hidden" id="option" name="option" value="add"> </TD>
            </TR>
          </TBODY>
        </TABLE>
</FORM>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td class="header">Sunsports &amp; Recreation</td>
        </tr>
        <tr> 
          <td>Office: (65) 63448344 <br>
            Fax: (65) 64844963 / 64752247<br>
            Adr:50 Jln Tembusu Singapore 438236 </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td class="header">Managing Director</td>
        </tr>
        <tr> 
          <td>
            Peter Koh<br>
            Mobile: (65) 91073488 / 97680275</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td class="header">Operations Manager</td>
        </tr>
        <tr> 
          <td>Kevin<br>
            Mobile: (65) 90887413 / 90040812</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td class="header">Chief Referee</td>
        </tr>
        <tr> 
          <td>Nasir<br>
            Mobile: (65) 91834350</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td class="header">Groundsman</td>
        </tr>
        <tr> 
          <td>Muti <br>
            Mobile: (65) 91065030</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td class="header">League Secretary </td>
        </tr>
        <tr> 
          <td>Connie<br>
            Office: 62919566 <br>
            Mobile: (65) 97903757</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td class="header">Marketing Rep</td>
        </tr>
        <tr> 
          <td>Rachel<br>
            Mobile: (65) 97951548</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td class="header">Youth Development </td>
        </tr>
        <tr> 
          <td>Thomas<br>
            Mobile: (65) 94559904</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td class="header">Marketing Mgt </td>
        </tr>
        <tr> 
          <td>Kevin<br>
            Mobile: (65) 90481234</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>	</td>
    <td width="20">&nbsp;</td>
    <td width="181" valign="top"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td><table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Events_all.asp"><b><font color="#1F4A71">SUNSPORTS'S 
                  NEWS</font></b></a></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <%
sql="select * from events order by eid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if not ( rs.eof or rs.bof ) then 
do
%>
              <tr> 
                <td width="5" valign="top"><img src="images/sun_l_sp.gif" width="5" height="10" border="0"></td>
                <td><a href="events_show.asp?eid=<%=rs("eid")%>" class="b"><%=rs("name")%></a> 
                </td>
              </tr>
              <%
rs.movenext
loop while not rs.eof 
end if
rs.close
set rs=nothing
%>
            </table>
            <br> <br> <br> <br> </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->