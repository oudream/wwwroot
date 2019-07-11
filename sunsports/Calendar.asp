<!--#include file="Top_sun.asp"-->
<%
user_uid=session("user_adminlevel")
if user_uid="" then
user_uid=""
elseif user_uid=4 then
user_uid=""
end if
field_id=request("fid")
if field_id="" then field_id=1
%>
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181" rowspan="2" valign="top"><table width="100%" border="0" cellpadding="3" cellspacing="0">
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
                <td><a href="TeamProfile.asp"><b><font color="#1F4A71">CHOOSE 
                  A FIELD</font></b></a></td>
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
                <td><a href="Calendar.asp?fid=<%=rs("fid")%>" title="<%=rs("name")%>" class="b"><%=rs("name")%></a></td>
              </tr>
              <%
rs.movenext
loop
rs.close
set rs=nothing
sql="select * from field where fid="&field_id
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
if not (rs.eof or rs.bof) then
fieldname=rs("name")
end if
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
          <td><font color="1F4A71">
		    #Slot 1 0900 - 1100 <br>
            #Slot 2 1100 - 1300 <br>
            #Slot 3 1300 - 1500 <br>
            #Slot 4 1500 - 1700 <br>
            #Slot 5 1700 - 1900 <br>
            #Slot 6 1900 - 2100 <br>
            #Slot 7 2100 - 2300 </font><br></td>
        </tr>
      </table></td>
    <td width="10" rowspan="2" valign="top">&nbsp;</td>
    <td> <form name="form1" method="get" action="Calendar.asp">
        <div align="center">[<font color="#F70000">#</font>&nbsp;<b>The red number 
          instead of today&nbsp;&nbsp;-&nbsp;&nbsp;<%=date%></b>]<br>
        </div>
        <div align="center"> 
          <select name="fid" id="fid" class="display_dropdown">
            <option value="1" selected>Select a field ...</option>
<%
sql="select * from field" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
do while not rs.eof 
%>
            <option value="<%=rs("fid")%>"><%=rs("name")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%> 
          </select>
          -- 
          <select name="month" class="display_dropdown">
            <option value="<%=month(date())%>" selected>Select the Month ...</option>
            <option>------------------</option>
            <option value="1">1</option>
            <option value="2">2</option>
            <option value="3">3</option>
            <option value="4">4</option>
            <option value="5">5</option>
            <option value="6">6</option>
            <option value="7">7</option>
            <option value="8">8</option>
            <option value="9">9</option>
            <option value="10">10</option>
            <option value="11">11</option>
            <option value="12">12</option>
          </select>
          -- 
          <select name="year" class="display_dropdown">
            <option value="2004" selected>Select the Year ...</option>
            <option>------------------</option>
            <option value="2001">2001</option>
            <option value="2002">2002</option>
            <option value="2003">2003</option>
            <option value="2004">2004</option>
            <option value="2005">2005</option>
            <option value="2006">2006</option>
            <option value="2007">2007</option>
            <option value="2008">2008</option>
            <option value="2009">2009</option>
            <option value="2010">2010</option>
          </select>
          --&gt; 
          <input type="submit" name="Submit" value="Display" class="button">
          <font face="verdana" size="2"></font> </div>
      </form></td>
  </tr>
  <tr> 
    <td>
<%
navday=request("day")
navmonth = request("month")
navyear = request("year")
navfield = request("fid")

if navfield = "" then
	navfield = 1
end if 

if navfield=5 then
	ai=7
else
	ai=5
end if

If navmonth = "" Then
	navmonth = Month(Date)
End If

If navyear = "" Then
	navyear = Year(Date)
End If

If navday = "" Then
	navday = day(Date)
End If



firstday = Weekday(CDate(cstr(navmonth) & "/" & cstr(1) & "/" & cstr(navyear)))

'！！！！！！！！！！！！！！！！！！登僅云埖葎謹富爺！！！！！！！！！！！！！！！！！！！！！！！！

leapTestNumbers = navyear / 4
leapTest = (leapTestNumbers) - Round(leapTestNumbers)
If navmonth = 2 Then
	If leapTest <> 0 Then
		lastDate = 28
	Else
		lastDate = 29
	End If
ElseIf ((navmonth = 4) OR (navmonth = 6) OR (navmonth = 9) OR (navmonth = 11)) Then
	lastDate = 30
Else
	lastDate = 31
End If

'！！！！！！！！！！！！！！！！！！登僅貧匯匈��和匯匈！！！！！！！！！！！！！！！！！！！！！！！！
lastMonth = navmonth - 1
lastYear = navyear
If lastMonth < 1 Then
	lastMonth = 12
	lastYear = lastYear - 1
End If

nextMonth = navmonth + 1
nextYear = navyear
If nextMonth >12 Then
	nextMonth = 1
	nextYear = nextYear + 1
End If


dateCounter = 1
weekCount = 1

DateEnd = lastDate
DateBegin = firstDate

if isempty(zdate) then zdate=navmonth&"-"&datecounter&"-"&navyear
%> <table  width="99%" height="160" border="0" cellpadding="5" cellspacing="1" bgcolor="#993300"  id="AutoNumber3" style="border-collapse: collapse"  >
        <tr bgcolor="#FF9900"> 
          <td height="17" colspan="7" align="left" valign="top" > <table width="100%" height="10" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="20">
                    <a href="Calendar.asp?month=<%=lastMonth%>&year=<%=lastYear%>&day=<%=dateCounter%>&date=<%=lastmonth&"/"&dateCounter&"/"&lastyear%>&fid=<%=navfield%>""> 
                    <img src="images/p_left.gif" width="13" height="20" border=0></a></td>
                <td width="80" align="center"> <%=navyear%>-<%=navMonth%></td>
                <td width="99">
                    <a href="Calendar.asp?month=<%=nextMonth%>&year=<%=nextYear%>&day=<%=dateCounter%>&date=<%=nextMonth&"/"&dateCounter&"/"&nextYear%>&fid=<%=navfield%>""> 
                    <img src="images/p_right.gif" width="13" height="20" border=0></a></td>
                <td width="368" valign="middle"><strong><%=fieldname%></strong>&nbsp;</td>
              </tr>
            </table></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="24" align="center" > <font color="#000000"><b><font size="2"> 
            SUN </font></b></font></td>
          <td height="24" align="center" bgcolor="#FFFFFF" > <font color="#000000"><b><font size="2"> 
            MON </font></b></font></td>
          <td height="24" align="center" > <font color="#000000"><b><font size="2"> 
            TUES</font></b></font></td>
          <td height="24" align="center" > <font color="#000000"><b><font size="2"> 
            WED</font></b></font></td>
          <td height="24" align="center" > <font color="#000000"><b><font size="2"> 
            THURS</font></b></font></td>
          <td height="24" align="center" > <font color="#000000"><b><font size="2"> 
            FRI</font></b></font></td>
          <td height="24" align="center" > <font color="#000000"><b><font size="2"> 
            SAT</font></b></font></td>
        </tr>
        <tr> 
          <% dim AArray() %>
          <% Do while weekCount <= 7 dateSelect = navmonth & "/" & dateCounter & "/" & navyear %>
          <% If weekCount < firstDay Then %>
          <td bgcolor="#FFFFFFF"></td>
          <% else %>
          <td bgcolor="#FFFFFF"> 
            <!-- start -->
<%
Set rss = Server.CreateObject("ADODB.RecordSet")
rss.Open "SELECT * FROM main where fid="&navfield&" and "&"year="&navyear&" and "&"month="&navmonth&" and "&"day="&dateCounter&" order by quantum", conn, 1, 1
zrecordcount=rss.recordcount
rss.cachesize=4
%> <%
redim AArray(ai)
j=1
k=1
if zrecordcount<>0 then
for i=1 to ai
	if rss("quantum")=j then
	AArray(j)="<font color=FF0000>"+cstr(j)+". Booked"+"</font><br>"
		if zrecordcount=k then
			rss.movefirst
		else
			rss.movenext
		end if
	k=k+1
	else
	AArray(j)=cstr(j)+". Available"+"<br>"
	end if 
j=j+1
next
else
j=1
for i=1 to ai
	AArray(j)=cstr(j)+". Available"+"<br>"
j=j+1
next
end if
%> 
            <!-- stop -->
            <%if clng(navyear)=clng(year(date)) and clng(navmonth)=clng(month(date)) and clng(datecounter)=clng(day(date)) then
	  %> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td align="right"><font color="#FF0000"><b><%=dateCounter%></b></font></td>
              </tr>
            </table>
            <% 
j=1 
for i=1 to ai 
if len(AArray(j))>23 then
response.Write(AArray(j))
else
if user_uid="" then
response.Write(AArray(j)) 
else
response.Write("<a href=emailtobuy.asp?fid="&navfield&"&zdate="&zdate&"&quantum="&j&" class=c target=_blank>"&AArray(j)&"</a>") 
end if
end if
j=j+1 
next 
%>
            <%elseif dateCounter=navday then%> 
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td align="right"><b><%=dateCounter%></b></td>
              </tr>
            </table>
            <% 
j=1 
for i=1 to ai 
if len(AArray(j))>23 then
response.Write(AArray(j))
else
if user_uid="" then
response.Write(AArray(j)) 
else
response.Write("<a href=emailtobuy.asp?fid="&navfield&"&zdate="&zdate&"&quantum="&j&" class=c target=_blank>"&AArray(j)&"</a>") 
end if
end if
j=j+1 
next 
%>
            <%else%> 
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td align="right"><b><%=dateCounter%></b></td>
              </tr>
            </table>
            <% 
j=1 
for i=1 to ai 
if len(AArray(j))>23 then
response.Write(AArray(j))
else
if user_uid="" then
response.Write(AArray(j)) 
else
response.Write("<a href=emailtobuy.asp?fid="&navfield&"&zdate="&zdate&"&quantum="&j&" class=c target=_blank>"&AArray(j)&"</a>") 
end if
end if
j=j+1 
next 
%>
            <%end if%> <%
erase AArray
rss.close
%> </td>
          <%
dateCounter = dateCounter + 1
zdate=navmonth&"-"&datecounter&"-"&navyear
end if

weekCount = weekCount + 1
Loop
weekCount = 1
%>
        </tr>
        <% 
dim BArray() 
%>
        <% Do while dateCounter <= lastDate %>
        <tr> 
          <% Do while weekCount <= 7 dateSelect = navmonth & "/" & dateCounter & "/" & navyear %>
          <% If dateCounter > lastDate Then %>
          <td bgcolor="#FFFFFF"></td>
          <% else %>
          <td bgcolor="#FFFFFF"> 
            <!-- start -->
            <%
Set rss = Server.CreateObject("ADODB.RecordSet") 
rss.Open "SELECT * FROM main where fid="&navfield&" and "&"year="&navyear&" and "&"month="&navmonth&" and "&"day="&dateCounter&" order by quantum", conn, 1, 1
zrecordcount=rss.recordcount
rss.cachesize=4
%> <%
redim BArray(ai)
j=1
k=1
if zrecordcount<>0 then
for i=1 to ai
	if rss("quantum")=j then
	BArray(j)="<font color=FF0000>"+cstr(j)+". Booked"+"</font><br>"
		if zrecordcount=k then
			rss.movefirst
		else
			rss.movenext	
		end if
	k=k+1
	else
	BArray(j)=cstr(j)+". Available"+"<br>"
	end if 
j=j+1
next
else
j=1
for i=1 to ai
	BArray(j)=cstr(j)+". Available"+"<br>"
j=j+1
next
end if
%> 
            <!-- stop -->
            <%if clng(navyear)=clng(year(date)) and clng(navmonth)=clng(month(date)) and clng(datecounter)=clng(day(date)) then
	  %> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td align="right"><font color="#FF0000"><b><%=dateCounter%></b></font></td>
              </tr>
            </table>
            <%
j=1
for i=1 to ai 
if len(BArray(j))>23 then
response.Write(BArray(j))
else
if user_uid="" then
response.Write(BArray(j)) 
else
response.Write("<a href=emailtobuy.asp?fid="&navfield&"&zdate="&zdate&"&quantum="&j&" class=c target=_blank>"&BArray(j)&"</a>") 
end if
end if
j=j+1 
next 
%>
             <%else%> <%if dateCounter=navday then%> 
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td align="right"><b><%=dateCounter%></b></td>
              </tr>
            </table>
            <%
j=1
for i=1 to ai 
if len(BArray(j))>23 then
response.Write(BArray(j))
else
if user_uid="" then
response.Write(BArray(j)) 
else
response.Write("<a href=emailtobuy.asp?fid="&navfield&"&zdate="&zdate&"&quantum="&j&" class=c target=_blank>"&BArray(j)&"</a>") 
end if
end if
j=j+1 
next 
%>
            <%else%>
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td align="right"><b><%=dateCounter%></b></td>
              </tr>
            </table>
            <%
j=1
for i=1 to ai 
if len(BArray(j))>23 then
response.Write(BArray(j))
else
if user_uid="" then
response.Write(BArray(j)) 
else
response.Write("<a href=emailtobuy.asp?fid="&navfield&"&zdate="&zdate&"&quantum="&j&" class=c target=_blank>"&BArray(j)&"</a>") 
end if
end if
j=j+1 
next 
%>
            <%end if%> <%end if%> <%
erase AArray

rss.close
%> </td>
          <% 
dateCounter = dateCounter + 1
zdate=navmonth&"-"&datecounter&"-"&navyear
end if

weekCount = weekCount + 1
Loop
weekCount = 1
%>
        </tr>
        <% Loop %>
      </table>
      <!-- 霞編何蛍  --start -->
      <%
'response.Write(zrecordcount)&"<br>"
'response.Write(navyear)&"/"
'response.Write(navmonth)&"/"
'response.Write(navday)&"<br>"
'response.Write(navfield)
%> 
      <!-- 霞編何蛍  --start -->
      <%=wwwww%> </td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->