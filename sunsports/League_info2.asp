<SCRIPT LANGUAGE="JavaScript">
<!--

function getcidweek(targ,cid,selObj,restore){ //v3.0
  eval(targ+".location='league.asp?cid="+cid+"&mweek="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

//-->

</SCRIPT>

<!--#include file="Top_sun.asp"-->
<table width="776" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181" valign="top"> <br>
      <table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td> 
            <!-- ========================================================================================================		  

													cup table ============star

 ======================================================================================================== -->
            <B></B> 
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr>
                <td width="158"><a href="League.asp"><b><font color="#1F4A71">CHOOSE A LEAGUE 
                  </font></b></a></td>
              </tr>
            </table> 
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <%
sql="select * from tournament where type='league' order by id_order" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
do while not rs.eof
%>
              <tr> 
                <td width="5" valign="top">
                  <img src="images/sun_l_sp.gif" width="5" height="10" border="0"> 
                  </td>
                <td><a href="League.asp?cid=<%=rs("cid")%>" title="<%=rs("name")%>" class="b"><%=rs("name")%></a></td>
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
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Cup.asp"><b><font color="#1F4A71">CUP MATCHES 
                  </font></b></a><font color="#1F4A71">&nbsp;+ +</font></td>
              </tr>
            </table> 
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Friendlies.asp"><b><font color="#1F4A71">FRIENDLY MATCHES 
                  </font></b></a></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="TeamProfile.asp"><b><font color="#1F4A71">TEAMPORFILE</font></b></a> 
                  <font color="#1F4A71">&nbsp;+ +</font> </td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="League_info.asp"><b><font color="#1F4A71">LEAGUE INFORMATION</font></b></a></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="League_rules.asp"><b><font color="#1F4A71">LEAGUE 
                  MATCH RULES</font></b></a></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
    </td>
    <td width="579" valign="top"><strong><br>
      Fantastic Prizes</strong> <br>
      <br>
      <p><strong>Stage I:</strong> - SunSport Super League V(Champion, Super Sunday, 
        Challenge&amp; Biz Co League). Product/ cash voucher will be awarded as 
        follow.<br>
        Champion: - $1000 cash or product voucher &amp; trophy <br>
        1st Runner-up - $300 cash or product voucher &amp; trophy<br>
        2 Runner-up: - $200 cash or product voucher</p>
      <p><strong>Stage II:</strong> - Sunsport Cup Winner Cup V. Product/ cash 
        voucher will be awarded as follow.<br>
        Champion: - $1000 cash or product voucher &amp; trophy<br>
        1st Runner-up: - $300 cash or product voucher &amp; trophy<br>
        2 Runner-up: - $200 cash or product voucher</p>
      <p><br>
        <strong>Stage IV:</strong>- SunSport Biz &amp; Corporation League (Cup 
        Winner Cup). Product/ cash voucher will be awarded as follow.<br>
        Champion :- $400 cash or product voucher &amp; trophy <br>
        1st Runner-up :- $200 cash or product voucher &amp; trophy<br>
        2 Runner-up :- $100 cash or product voucher <br>
        Golden Boots Award :- $100 product voucher<br>
        Fair Play Team Award: $100 product voucher<br>
        Best Referee Award : $100 product voucher</p>
      <p><strong><br>
        Sunsport and Recreation Super League V 2003/04 Itinerary</strong></p>
      <p><font color="#FF0000">Date: Mar 2003- Mar 2004 (34 games)</font><br>
        Stage 1: SunSport Super League (Champion, Super Sunday &amp; Challenge 
        League). </p>
      <p><font color="#FF0000">Date: May 2003</font><br>
        Award Presentation, Sport Club Membership Launch &amp; Media Coverage 
        <br>
        For 2002 Super Leagues, SunSport Recreation Sport Club &amp; Sentosa Nite 
        Field</p>
      <p><font color="#FF0000">Date: Jan 2004-Mar 2004 (5-8 games)</font><br>
        Stage 2: SunSport Cup Winner Cup V. </p>
      <p><font color="#FF0000">Date: June 2003-Aug 2003 </font><br>
        Stage 3: Special Edition SunSport Sentosa 7 A-side Cup (Late June-Aug)</p>
      <p><font color="#FF0000">Date:Aug/Sept 2003 (1 day event)</font><br>
        Stage 4: SunSport Community Charity Shield</p>
      <p><font color="#FF0000">Date: May 2004</font><br>
        Award Presentation &amp; Media Coverage for 2003/04 Super League. </p>
      <p>All in all we will play 40-42 games /weeks of soccer for the season.<br>
      </p>
      <strong> </strong></td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->