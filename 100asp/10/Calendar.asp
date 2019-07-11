<%@ Language=VBScript %>
<Html>
<Title>
小日历
</title>
<body>
<%
Function CountDays(iMonth,iYear)
    Select Case iMonth
    case 1,3,5,7,8,10,12
	CountDays=31
    case 2
	if IsDate("2/29/" & iYear) Then
	    CountDays=29
	else
	    CountDays=28
        end if
    case 4,6,9,11
        CountDays=30
    End Select
End Function

Function FirstDay(iMonth,iYear)
    FirstDay=WeekDay(iMonth & "/1/" & iYear)
End Function

dim mMonth,mYear
mMonth=Month(Date())
mYear=Year(Date())
mDate=Day(Date())
response.write "<center>" & mYear & "年" & mMonth & "月" & "</center><hr>"
%>
<table border=1 align=center><tr>
<td align=right>星期日</td>
<td align=right>星期一</td>
<td align=right>星期二</td>
<td align=right>星期三</td>
<td align=right>星期四</td>
<td align=right>星期五</td>
<td align=right>星期六</td>
</tr><tr>
<%
j=1
    for i=1 to 42
        response.write "<td align=right>"
        if i>=FirstDay(mMonth,mYear) and j<=CountDays(mMonth,mYear) then
            if mDate=j then
		response.write "<font color=blue>" & j & " </font>"
	    else
		response.write j
	    end if
	    j=j+1
        else
	    response.write   " &nbsp; "
        end if
	response.write "</td>"
    	if i mod 7=0 then
	    response.write "</tr><tr>"
        end if
    next 
    
%>
</tr></table>
</body>
</html>
