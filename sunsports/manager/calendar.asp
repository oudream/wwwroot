<!--#include file="../conn/conn.asp"-->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body>

<form name="form1" method="get" action="calendar.asp">
                    
  <div align="center"><font color="#F70000">■</font>Today<br>
      </div> 
  <div align="center"> 
    <select name="field" id="field">
      <option value="1" selected>Select a field ...</option>
      <option value="1">Katong A</option>
      <option value="2">Katong B</option>
      <option value="3">New Town</option>
      <option value="4">Saint Andrew's</option>
      <option value="5">Sentosa East</option>
      <option value="6">Sentosa West Field</option>
      <option value="7">Serangoon</option>
      <option value="8">Tjc</option>
      <option value="9">Unknown</option>
    </select>
    --
<select name="month" >
      <option value="1" selected>Select the Month ...</option>
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
    <select name="year" >
      <option value="2004" selected>Select the Year ...</option>
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
    <input type="submit" name="Submit" value="Display" >
  </div>
</form>

<%
navday=request("day")
navmonth = request("month")
navyear = request("year")
navfield = request("field")

if navfield = "" then
	navfield = "1"
end if 

if trim(cstr(navfield))=cstr(5) then
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

'——————————————————判断本月为多少天————————————————————————

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

'——————————————————判断上一页，下一页————————————————————————
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


%>


	<table border="1" cellpadding="0" cellspacing="1" style="border-collapse: collapse"  id="AutoNumber3"  width="100%" height="160" bordercolor="#666666"  >
		<tr align="left" valign="top" bgcolor="#95B8FF"> 
			<td height="17" colspan="7" > 
			<table border="0" cellspacing="0" width="100%" height="10">
				<tr valign="middle"> 
					<td width="21%" height="15"> <div align="right"><font face="verdana" size="2"><b> 
					<a href="calendar.asp?month=<%=lastMonth%>&year=<%=lastYear%>&day=<%=dateCounter%>&date=<%=lastmonth&"/"&dateCounter&"/"&lastyear%>&field=<%=navfield%>""> 
					<img src="../images/cl03.gif" border=0> </a></b></font></div></td>
					<td width="58%" height="15"> <div align="center"> 
					<%=navfield%>--<%=navyear%>--<%=navMonth%>--<font face="verdana" size="4"><b> 
					&nbsp;</b></font></div></td>
					<td width="21%" height="15"> <div align="left"><font face="verdana" size="2"><b> 
					<a href="calendar.asp?month=<%=nextMonth%>&year=<%=nextYear%>&day=<%=dateCounter%>&date=<%=nextMonth&"/"&dateCounter&"/"&nextYear%>&field=<%=navfield%>""> 
					<img src="../images/cl04.gif" border=0> </a></b></font></div></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr align="center" valign="middle" bgcolor="#FFFFFF"> 
			<td height="24" > <div align="center"><font color="#000000"><b><font size="2"> 
        SUN </font></b></font></div></td>
			<td height="24" > <div align="center"><font color="#000000"><b><font size="2"> 
        MON </font></b></font></div></td>
			<td height="24" > <div align="center"><font color="#000000"><b><font size="2"> 
        TUES</font></b></font></div></td>
			<td height="24" > <div align="center"><font color="#000000"><b><font size="2"> 
        WED</font></b></font></div></td>
			<td height="24" > <div align="center"><font color="#000000"><b><font size="2"> 
        THURS</font></b></font></div></td>
			<td height="24" > <div align="center"><font color="#000000"><b><font size="2"> 
        FRI</font></b></font></div></td>
			<td height="24" > <div align="center"><font color="#000000"><b><font size="2"> 
        SAT</font></b></font></div></td>
		</tr>

		<tr> 
<% dim AArray() %>
<% Do while weekCount <= 7 dateSelect = navmonth & "/" & dateCounter & "/" & navyear %>
<% If weekCount < firstDay Then %>
			<td bgcolor="#FFFFFFF"></td>
<% else %>
			<td bgcolor="#FFFFFF" valign="middle" align="center">
<!-- start -->
<%
Set rss = Server.CreateObject("ADODB.RecordSet")
rss.Open "SELECT * FROM main where field="&"'"&cstr(navfield)&"'"&" and "&"year="&"'"&cstr(navyear)&"'"&" and "&"month="&"'"&cstr(navmonth)&"'"&" and "&"day="&"'"&cstr(dateCounter)&"'"&" order by quantum", conn, 1, 1
zrecordcount=rss.recordcount
rss.cachesize=4
%>
<%
redim AArray(ai)
j=1
k=1
if zrecordcount<>0 then
for i=1 to ai
	if trim(cstr(rss("quantum")))=cstr(j) then
	AArray(j)=cstr(j)+". Booked"+"<br>"
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
<%if clng(navyear)=year(date) and clng(navmonth)=month(date) and datecounter=day(date) then%>
			<font color="#CC0000"><b><br><%=dateCounter%></b><br>
<% 
j=1 
for i=1 to ai 
response.Write(AArray(j)) 
j=j+1 
next 
%>
			</font> 
<%elseif dateCounter=clng(navday) then%>
			<font color="#0000FF"><b><br><%=dateCounter%></b><br>
<% 
j=1 
for i=1 to ai 
response.Write(AArray(j)) 
j=j+1 
next 
%>
			</font> 
<%else%>
			<font color="#000000"><b><br><%=dateCounter%></b><br>
<% 
j=1 
for i=1 to ai 
response.Write(AArray(j)) 
j=j+1 
next 
%>
			</font> 
<%end if%> 
<%
erase AArray
rss.close
%>
			</td>
<% 
dateCounter = dateCounter + 1
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
                      <td bgcolor="#FFFFFF" valign="middle" align="center">

<!-- start -->
<%
Set rss = Server.CreateObject("ADODB.RecordSet") 
rss.Open "SELECT * FROM main where field="&"'"&cstr(navfield)&"'"&" and "&"year="&"'"&cstr(navyear)&"'"&" and "&"month="&"'"&cstr(navmonth)&"'"&" and "&"day="&"'"&cstr(dateCounter)&"'"&" order by quantum", conn, 1, 1
zrecordcount=rss.recordcount
rss.cachesize=4
%>
<%
redim BArray(ai)
j=1
k=1
if zrecordcount<>0 then
for i=1 to ai
	if trim(cstr(rss("quantum")))=cstr(j) then
	BArray(j)=cstr(j)+". Booked"+"<br>"
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

                        <%if clng(navyear)=year(date) and clng(navmonth)=month(date) and datecounter=day(date) then%>
                        <font color="#CC0000"><b><%=dateCounter%></b><br>
<%
j=1
for i=1 to ai 
response.Write(BArray(j)) 
j=j+1 
next 
%>
								</font> 
                        <%else%>
                        <%if dateCounter=clng(navday) then%>
                        <font color="#0000FF"><b><%=dateCounter%></b><br>
<%
j=1
for i=1 to ai 
response.Write(BArray(j)) 
j=j+1 
next 
%>
								</font> 
                        <%else%>
                        <font color="#000000"><b><%=dateCounter%></b><br>
<%
j=1
for i=1 to ai 
response.Write(BArray(j)) 
j=j+1 
next 
%>
								</font> 
                        <%end if%>
                        <%end if%>
<%
erase AArray

rss.close
%>
								</td>
                      <% 
dateCounter = dateCounter + 1
end if

weekCount = weekCount + 1
Loop
weekCount = 1
%>
                    </tr>
                    <% Loop %>
                  </table> 
                  
<!-- 测试部分  --start -->
<%
'response.Write(zrecordcount)&"<br>"
'response.Write(navyear)&"/"
'response.Write(navmonth)&"/"
'response.Write(navday)&"<br>"
'response.Write(navfield)
%>
<!-- 测试部分  --start -->
</body>
</html>
















