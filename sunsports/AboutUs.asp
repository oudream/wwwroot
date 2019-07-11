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
            <table width="169" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td width="163"><a href="TeamProfile.asp"><b><font color="#1F4A71">TEAMPORFILE</font></b></a> 
                  <font color="#1F4A71">&nbsp;+ +</font> </td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="TeamProfile.asp"><b><font color="#1F4A71">FIELDS 
                  BOOKING</font></b></a></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <%
sql="select * from field" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
do while not rs.eof
%>
              <tr> 
                <td width="5" valign="top"> <img src="images/sun_l_sp.gif" width="5" height="10" border="0"> 
                </td>
                <td><a href="Calendar.asp?field=<%=rs("fid")%>" title="<%=rs("name")%>" class="b"><%=rs("name")%></a></td>
              </tr>
              <%
rs.movenext
loop
rs.close
set rs=nothing
%>
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
    <td width="20">&nbsp; </td>
    <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="27">&nbsp;</td>
        </tr>
        <tr> 
          <td height="50" valign="top" class="header_b">ABUOT SUNSPORTS</td>
        </tr>
        <tr> 
          <td class="header_a">Sunsports &amp; Recreation</td>
        </tr>
        <tr> 
          <td>Sunsports is a private events organiser for sports related activities 
            that caters to individuals, companies, MNC's, schools, institutions 
            and government agencies. We actively encourage participation in sporting 
            activities for the joint benefit of promoting a healthy lifestyle 
            as well as team co-operation through fun sports.<br>
            <br>
          </td>
        </tr>
        <tr> 
          <td class="header_a">Include</td>
        </tr>
        <tr> 
          <td>Organised sporting events offered include: </td>
        </tr>
        <tr>
          <td>
            <LI>Soccer <br>
            <LI>Ball Games <br>
            <LI>X country <br>
            <LI>Island Cycling <br>
            <LI>Water Sports <br>		  </td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td class="header_a">Matches</td>
        </tr>
        <tr>
          <td>As Soccer is the most popular sports , Sunsports organise many soccer 
            matches here for years.We have league maches every weekend and we 
            have &quot;SunSports Cup Series V&quot; and &quot;Sunsports Sentosa 
            Cup (7 A-side)&quot; every year.</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td class="header_a">Fields</td>
        </tr>
        <tr> 
          <td>Sunsports has Katong A ,Katong B , New Town , Saint Andrew's ,<br>
            Sentosa East , Sentosa West Field , Serangoon and Tjc for games or 
            matches.If you want more infomation please <a href="Fields.asp">click 
            here</a> . If you want to book fields for games or matches , please 
            <a href="Contact.asp">click here</a> to view our contact way.Thank 
            you! <br></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table>
      
    </td>
    <td width="20">&nbsp;</td>
    <td width="181" valign="top">&nbsp;</td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->