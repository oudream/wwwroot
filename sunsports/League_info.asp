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
      Soccer League Information</strong><br>
      <br>
      Sunsport and Recreation is proud to invite you and your organisation to 
      participate in the soccer tournament known as the &#8220;SunSport Super 
      League/Cup 2003/2004, eleven a side soccer tournament. Namely 4 leagues 
      Champion, Super Sunday, SR Challenge &amp; Biz Co-operation Nite League, 
      18 teams will be invited to participate in each league. <br>
      St. Andrew JC &amp; St.Wilf or otherwise determine by the organising committee. 
      <p><br>
        All matches will be played at Santos, Serangoon Stadium, Newtown, Katong 
        (Haig Rd), </p>
      <p><strong><br>
        Breakdown to Stages</strong></p>
      <p><strong>Stage 1:</strong> SunSport Super League (Champion, Super Sunday, 
        Challenge &amp; Biz Co League) <br>
        Each league would consist of 18 teams (namely Champion, Super Sunday &amp; 
        Challenge League) playing 17x2=34 games. Top teams of each group will 
        play in the 1 day Family Carnival Event at nominal fees. </p>
      <p>Champion League would be played (2/3 fixtures) Sun 3-7pm &amp; If possible 
        (1/3 fixtures) Sat 3-7pm plus Sun 9-1pm.<br>
        Super Sunday League would be played Sunday 9-1pm<br>
        Challenge League would be played Sat 3-7 &amp; Sun 3-7pm<br>
        Biz Co-operation Nite League would be played wkday 7-11pm <br>
        Duration 40min per half. Expected of Completion: 34 wks </p>
      <p><strong>Stage 2:</strong> SunSport Super Cup (Cup Winner Cup).<br>
        The top 9 teams of each League (total of 32 teams) would be selected to 
        play the knockout phase of Cup Winner Cup. Where all banking &amp; corporation 
        teams would be group into play. Each team would be matched according to 
        draw-lot. That will decide the Champion of Sunsport Super Cup. The Champion 
        would proceed to play in the Cup Winner Cup.<br>
        Duration 40 min per half. Expected date of Completion: 8 wks</p>
      <p><strong>Stage 3:</strong> Special Edition SunSport Sentosa Cup (June-Aug) 
        optional<br>
        All teams will play in the SunSport Sentosa Cup 7 A &#8211; side tournament. 
        Drawing lot where teams played exactly like the World Cup format with 
        some teams seeded in each group will do matching of teams. All teams will 
        play &amp; name according to the exact nation team in the World Cup.<br>
        Duration 15 min each half @ 30 min per game.<br>
        Detail out later. Expected date of Completion: 3-4wks </p>
      <p><strong>Stage 4:</strong> SunSport Community Charity Shield <br>
        All teams will then be invited to play in the Charity Cup game at the 
        end of the league games at a nominal fee. <br>
        Detail out later. Expected date of Completion: 1wk <br>
        All games in the stage 1 &amp; 2 will be played with referee, drink, captain 
        band, match-ball &amp; match report notice with lineman as optional. <a href="League_info2.asp">More</a><br>
      </p></td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->