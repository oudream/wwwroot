<!--#include file="conn/conn.asp" -->
<html>
<head>
<title>CONTACT , PJEANS.COM</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" rel=stylesheet 
type=text/css>
</head>

<body>
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
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=600>
  <TBODY>
    <TR> 
      <TD><IMG height=10 src="images/px.gif" width=10></TD>
    </TR>
    <TR> 
      <TD> <TABLE border=0 cellPadding=0 cellSpacing=0>
          <TBODY>
            <TR> 
              <TD align=left rowSpan=3 vAlign=top><img src="images/pjeans_logo.gif" width="210" height="70"></TD>
              <TD height="35" colSpan=2 align=right vAlign=bottom class=date> 
                <SPAN id=clock> 
                <SCRIPT language=JavaScript>

var dayarray=new Array("SUNDAY","MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY")
var montharray=new Array("JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER")
function getthedate(){
var mydate=new Date()
var year=mydate.getYear()
if (year < 1000)
year+=1900
var day=mydate.getDay()
var month=mydate.getMonth()
var daym=mydate.getDate()
if (daym<10)
daym="0"+daym
var hours=mydate.getHours()
var minutes=mydate.getMinutes()
var seconds=mydate.getSeconds()
var dn="AM"
if (hours>=12)
dn="PM"
if (hours>12){
hours=hours-12
}
{
 d = new Date();
 Time24H = new Date();
 Time24H.setTime(d.getTime() + (d.getTimezoneOffset()*60000) + 3600000);
 InternetTime = Math.round((Time24H.getHours()*60+Time24H.getMinutes()) / 1.44);
 if (InternetTime < 10) InternetTime = '00'+InternetTime;
 else if (InternetTime < 100) InternetTime = '0'+InternetTime;
}
if (hours==0)
hours=12
if (minutes<=9)
minutes="0"+minutes
if (seconds<=9)
seconds="0"+seconds

var cdate=dayarray[day] + "&nbsp;&nbsp;" + daym + "&nbsp;&nbsp;" +montharray[month]+ "&nbsp;&nbsp;" + hours + ":"+minutes+":"+seconds+" "+dn+" "
if (document.all)
document.all.clock.innerHTML=cdate
else if (document.getElementById)
document.getElementById("clock").innerHTML=cdate
else
document.write(cdate)
}
if (!document.all&&!document.getElementById)
getthedate()
function goforit(){
if (document.all||document.getElementById)
setInterval("getthedate()",1000)
}
window.onload=goforit
//  End -->
</SCRIPT>
                </SPAN> </TD>
            </TR>
            <TR> 
              <TD align=right colSpan=2>&nbsp; </TD>
            </TR>
            <TR> 
              <TD width=400 height="18" colSpan=2 align=right vAlign=bottom class=power><font color="#990000">Welcome 
                to Pjeans's Web Site</font></TD>
            </TR>
            <TR> 
              <TD colSpan=3><IMG height=1 src="images/px.gif" 
            width=1></TD>
            </TR>
            <TR> 
              <TD background=images/hpx.jpg colSpan=3><IMG height=2 
            src="images/px.gif" width=1></TD>
            </TR>
            <TR> 
              <TD colSpan=3><IMG height=10 src="images/px.gif" 
            width=1></TD>
            </TR>
            <TR> 
              <TD colSpan=3> <TABLE border=0 cellPadding=0 cellSpacing=0 width=600>
                  <TBODY>
                    <TR> 
                      <TD width=437><SPAN class=menu><a href="men.asp">MEN</a><IMG 
                  height=1 src="images/px.gif" width=10><IMG 
                  align=absMiddle height=15 src="images/hpx.gif" 
                  width=1><IMG height=1 src="images/px.gif" 
                  width=10></SPAN> <SPAN class=menu><a href="women.asp">WOMEN</a><IMG 
                  height=1 src="images/px.gif" width=10><IMG 
                  align=absMiddle height=15 src="images/hpx.gif" 
                  width=1><IMG height=1 src="images/px.gif" 
                  width=10></SPAN> <SPAN class=menu><a href="children.asp">CHILDREN</a><IMG 
                  height=1 src="images/px.gif" width=10><IMG 
                  align=absMiddle height=15 src="images/hpx.gif" 
                  width=1><IMG height=8 src="images/px.gif" 
                  width=10></SPAN> <SPAN class=menu><A 
                  href="about_us.asp">ABOUT US </A><IMG height=1 src="images/px.gif" 
                  width=10><IMG align=absMiddle height=15 
                  src="images/hpx.gif" width=1><IMG height=1 
                  src="images/px.gif" width=10></SPAN> <SPAN 
                  class=menu><A 
                  href="contact.asp">CONTACT US </A><IMG 
                  height=1 src="images/px.gif" width=10></SPAN> </TD>
                      <TD align=right 
      width=163></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
    <TR> 
      <TD> <TABLE border=0 cellPadding=0 cellSpacing=0>
          <TBODY>
            <TR> 
              <TD align=left vAlign=top><IMG height=15 
            src="images/px.gif" width=1></TD>
              <TD align=right><IMG height=5 src="images/px.gif" 
            width=1></TD>
            </TR>
            <TR> 
              <TD align=left colSpan=2 vAlign=top> <TABLE border=0 cellPadding=0 cellSpacing=0 width=600>
                  <TBODY>
                    <TR> 
                      <TD vAlign=top> <TABLE border=0 cellPadding=0 cellSpacing=0 width=100%>
                          <TBODY>
                            <TR> 
                              <TD class=header><P><B>Contact Us</B></P></TD>
                              <TD align=right class=icons><NOBR></NOBR></TD>
                            </TR>
                            <TR> 
                              <TD background=images/hpx.jpg 
                        colSpan=2><IMG height=2 
                        src="images/px.gif" width=1></TD>
                            </TR>
                            <TR> 
                              <TD colSpan=2><IMG height=6 
                        src="images/px.gif" width=1></TD>
                            </TR>
                            <TR align=left vAlign=top> 
                              <TD colSpan=2> <P><B><br>Please send us an <a href="mailto:sunsport@singnet.com.sg" class="b">email</a> 
                                            if you have any queries with regards 
                                            to participation in the various soccer 
                                            leagues. Alternatively, please use 
                                            the form below to submit any queries 
                                            or comments you may have.</B></P></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE border=0 cellPadding=0 cellSpacing=0 width=100%>
                          <TBODY>
                            <TR> 
                              <TD class=header width="43%">&nbsp;</TD>
                              <TD align=right class=icons width="57%">&nbsp;</TD>
                            </TR>
                            <TR> 
                              <TD background=images/hpx.jpg 
                        colSpan=2><IMG height=2 
                        src="images/px.gif" width=1></TD>
                            </TR>
                            <TR> 
                              <TD colSpan=2><IMG height=10 
                        src="images/px.gif" width=1></TD>
                            </TR>
                            <TR vAlign=top> 
                              <TD colSpan=2><table width="90%" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td> 
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
                                              <TD bgColor=#ffffff width="75%"> 
                                                <INPUT class=form id="name"
                                onblur="if(this.value=='')this.value='Name';" 
                                style="FONT-SIZE: 12px; WIDTH: 360px; background:white;border:1px solid #6184A3" onMouseOver="this.style.border='1px solid red'" onMouseOut="this.style.border='1px solid #6184A3'" 
                                onfocus="if(this.value=='Name')this.value='';" 
                                value=Name name=name
								
								> </TD>
                                            </TR>
                                            <TR bgColor=#ffffff> 
                                              <TD width="75%" height="27" bgColor=#ffffff> 
                                                <INPUT class=form id="email"
                                onblur="if(this.value=='')this.value='Email';" 
                                style="FONT-SIZE: 12px; WIDTH: 360px; background:white;border:1px solid #6184A3" onMouseOver="this.style.border='1px solid red'" onMouseOut="this.style.border='1px solid #6184A3'" 
                                onfocus="if(this.value=='Email')this.value='';" 
                                value=Email name=email> </TD>
                                            </TR>
                                            <TR bgColor=#ffffff> 
                                              <TD bgColor=#ffffff width="75%"> 
                                                <INPUT class=form id="subject"
                                onblur="if(this.value=='')this.value='Subject';" 
                                style="FONT-SIZE: 12px; WIDTH: 360px; background:white;border:1px solid #6184A3" onMouseOver="this.style.border='1px solid red'" onMouseOut="this.style.border='1px solid #6184A3'" 
                                onfocus="if(this.value=='Subject')this.value='';" 
                                value=Subject name=subject> </TD>
                                            </TR>
                                            <TR bgColor=#ffffff> 
                                              <TD bgColor=#ffffff width="75%"> 
                                                <TEXTAREA id="mailtext" class=form onblur="if(this.value=='')this.value='Mailtext';"                                 style="FONT-SIZE: 12px; WIDTH: 360px; background:white;border:1px solid #6184A3" onMouseOver="this.style.border='1px solid red'" onMouseOut="this.style.border='1px solid #6184A3'" 
 onfocus="if(this.value=='Mailtext')this.value='';" name=mailtext rows=10 wrap=VIRTUAL value="Mailtext">Mailtext</TEXTAREA> 
                                              </TD>
                                            </TR>
                                            <TR bgColor=#ffffff> 
                                              <TD bgColor=#ffffff> <input type=submit id="options" name="options" value="          [ Submit Here ]          " class="button"> 
                                                <input type="hidden" id="option" name="option" value="add"> 
                                              </TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                      </FORM>
                                      
                                    </td>
                                  </tr>
                                </table> </TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
    <TR> 
      <TD align=left class=menu vAlign=top> 
        <table width="75%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>Contact person: Peter Cheng<br>
              <br>
              Tel : 86-757-25593678<br>
              <br>
              Fax :86-757-25593623<br>
              <br>
              Mobile: 86-13380506250<br>
              <br>
              e-mail: peter@joyfulshow.com<br>
              <br>
              Address: No. 1 Hua’an RD. Jun’an town, Shunde, GuangDong, China</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table></TD>
    </TR>
    <TR> 
      <TD background=images/hpx.jpg width=414><IMG height=2 
            src="images/px.gif" width=1></TD>
    </TR>
    <TR> 
      <TD align=left class=menu vAlign=top><IMG height=3 
            src="images/px.gif" width=1></TD>
    </TR>
    <TR> 
      <TD align=left class=menu vAlign=top><SPAN class=menu><a href="men.asp">MEN</a><IMG 
            height=1 src="images/px.gif" width=10><IMG 
            align=absMiddle height=15 src="images/hpx.gif" 
            width=1><IMG height=1 src="images/px.gif" 
            width=10></SPAN> <SPAN class=menu><a href="women.asp">WOMEN</a><IMG 
            height=1 src="images/px.gif" width=10><IMG 
            align=absMiddle height=15 src="images/hpx.gif" 
            width=1><IMG height=1 src="images/px.gif" 
            width=10></SPAN> <SPAN class=menu><a href="children.asp">CHILDREN</a><IMG 
            height=1 src="images/px.gif" width=10><IMG 
            align=absMiddle height=15 src="images/hpx.gif" 
            width=1><IMG height=1 src="images/px.gif" 
            width=10></SPAN> <SPAN class=menu><A 
                  href="about_us.asp">ABOUT US</A><IMG height=1 src="images/px.gif" 
            width=10><IMG align=absMiddle height=15 
            src="images/hpx.gif" width=1><IMG height=1 
            src="images/px.gif" width=10></SPAN> <SPAN 
            class=menu><A 
                  href="contact.asp">CONTACT US</A><IMG 
            height=1 src="images/px.gif" width=10></SPAN> <br> <br>
        Copyright &copy; 2004 Pjeans.com , Inc. All rights reserved.<br>
        See <a href="contact.asp">Contact Us</a> for more information.</TD>
    </TR>
  </TBODY>
</TABLE>
</body>
</html>