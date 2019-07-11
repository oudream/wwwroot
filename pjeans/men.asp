<!--#include file="conn/conn.asp" -->
<html>
<head>
<title>MEN</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" rel=stylesheet 
type=text/css>
</head>

<body>
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
              <TD align=left vAlign=top>&nbsp;</TD>
              <TD>&nbsp;</TD>
            </TR>
            <TR> 
              <TD align=left colSpan=2 vAlign=top> <TABLE border=0 cellPadding=0 cellSpacing=0 width=600>
                  <TBODY>
                    <TR> 
                      <TD vAlign=top width=414> <TABLE border=0 cellPadding=0 cellSpacing=0 width=414>
                          <TBODY>
                            <TR> 
                              <TD class=header><p><b>MEN</b></p></TD>
                              <TD align=right class=icons><font color="#990000"><NOBR> 
                                <img src="images/dot-7.gif" width="8" height="10"> 
                                Products Guide.</NOBR></font></TD>
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
                              <TD height="240" colSpan=2><SPAN 
                        class=header_date><B>FASHION</B></SPAN><BR> <P><B>Here 
                                  you can write all your ideas down.EG.You can 
                                  write something about fashion,jeans,man,woman,child 
                                  and so on.ofcouse you can write your product 
                                  or your ...</B></P>
                                <P><font color="#990000">Summer Style for women 
                                  2004</font> What are the biggest trends we are 
                                  seeing teens take on? Out of all the clothes 
                                  we wear, pants are the most key. Your favorite 
                                  hip-hugging, soft-wearing, go-with-anything, 
                                  dress-up-or-down pants are an essential. Tight 
                                  tops and belly baring are okay. </P>
                                <p>Casual styles borrowed from sportswear, and 
                                  retro-style tops are a big trend. Karate styles, 
                                  with lots of black, white and red hues are popular. 
                                </p>
                                </TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <BR> <TABLE border=0 cellPadding=0 cellSpacing=0 width=414>
                          <TBODY>
                            <TR> 
                              <TD class=header width="43%">Pjeans.</TD>
                              <TD align=right class=icons width="57%"><NOBR> <A 
                        href="men_view.asp">New products</A></NOBR></TD>
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
                            <TR align=left vAlign=top> 
                              <TD colSpan=2> <%
dim men_l(2)
sql="select * from subs where sort_id=1110 or sort_id=1111  or sort_id=1112  or sort_id=1113 order by subs_id desc "
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
do while not rs.eof 
select case rs("sort_id")
case 1110 men_r="men/"+rs("pic")
case 1111 men_l(0)="men/"+rs("pic")
case 1112 men_l(1)="men/"+rs("pic")
case 1113 men_l(2)="men/"+rs("pic")
end select
rs.movenext
loop
rs.close
set rs=nothing
%> <TABLE border=0 cellPadding=0 cellSpacing=0 width=414>
                                  <TBODY>
                                    <TR align=left vAlign=top> 
                                      <TD width=131><IMG 
                              src="<%=men_l(0)%>" width="130" height="182" border=0></TD>
                                      <TD width=10>&nbsp;</TD>
                                      <TD width=131><img src="<%=men_l(1)%>" width="130" height="182"></TD>
                                      <TD width=10>&nbsp;</TD>
                                      <TD width=132><img src="<%=men_l(2)%>" width="130" height="182"></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <P></P></TD>
                      <TD width=20>&nbsp;</TD>
                      <TD width=166 vAlign=top> <TABLE border=0 cellPadding=0 cellSpacing=0 width=166>
                          <TBODY>
                            <TR> 
                              <TD width=8>&nbsp;</TD>
                              <TD align=right background=images/px.gif 
                      bgColor=#ffffff vAlign=bottom width=150>&nbsp;</TD>
                              <TD align=right width=8>&nbsp;</TD>
                            </TR>
                            <TR> 
                              <TD width=8>&nbsp;</TD>
                              <TD align=right background=images/px.gif 
                      bgColor=#ffffff class=power vAlign=bottom width=150> <IMG src="<%=men_r%>" width=150 height=227 
                        border=0></TD>
                              <TD align=right width=8>&nbsp;</TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <P align="center"><img src="images/r-w.gif" width="150" height="109"></P>
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td class=size10> New fashion statement makers will 
                              emerge as we approach the middle of the first decade 
                              of the 21st century. Global trends now emerge faster 
                              than ever before as the internet and text messaging, 
                              shortcuts transmission of ideas and concepts between 
                              individuals, bringing the next big thing to you 
                              and me in the space of seconds. </td>
                          </tr>
                        </table></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
    <TR> 
      <TD align=left class=menu vAlign=top><IMG height=10 
            src="images/px.gif" width=1></TD>
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