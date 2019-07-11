<%
Class Calendar
	Public Top
	Public Left
	Public Width
	Public Height
	Public Position
	Public ZIndex
	Public TitlebarColor
	Public TitlebarFont
	Public TitlebarFontColor
	Public TodayBGColor
	Public OnDayClick
	Public OnNextMonthClick
	Public OnPrevMonthClick
	Public ShowDateSelect
	Private mdDate
	Private msToday
	Private mnDay
	Private mnMonth
	Private mnYear
	Private mnDayMonthStarts
	Private mnDaysInMonth
	Private mcolDays
	Private mbDaysInitialized
	
	Private Sub Class_Initialize()
		Top = 0
		Left = 0
		Width = 500
		Height= 500
		Position = "absolute"
		TitlebarColor = "darkblue"
		TitlebarFont = "arial"
		TitlebarFontColor = "white"
		TodayBGColor = "skyblue"
		ShowDateSelect = True
		msToday =  FormatDateTime(DateSerial(Year(Now()), Month(Now()), Day(Now())), 2)
		zIndex = 1
		
		Set mcolDays = Server.CreateObject("Scripting.Dictionary")
		If Request("date") <> "" Then SetDate(Request("date")) Else SetDate(Now())

		OnDayClick = Request.ServerVariables("SCRIPT_NAME")
		OnNextMonthClick = Request.ServerVariables("SCRIPT_NAME") & "?date=" & Server.URLEncode(DateSerial(mnYear, mnMonth + 1, mnDay))
		OnPrevMonthClick = Request.ServerVariables("SCRIPT_NAME") & "?date=" & Server.URLEncode(DateSerial(mnYear, mnMonth - 1, mnDay))

		mbDaysInitialized = False
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(mcolDays) Then
			mcolDays.RemoveAll
			Set mcolDays = Nothing
		End If
	End Sub
	
	Public Property Get GetDate()
		GetDate = mdDate
	End Property
	
	Public Property Get DaysInMonth()
		DaysInMonth = mnDaysInMonth
	End Property
	
	Public Property Get WeeksInMonth()
		If (mnDayMonthStarts + mnDaysInMonth - 1) > 35 Then
			WeeksInMonth = 6
		Else
			WeeksInMonth = 5
		End If
	End Property
	
	Public Property Get Days(nIndex)
		If Not mbDaysInitialized Then InitDays()
		If mcolDays.Exists(nIndex) Then Set Days = mcolDays.Item(nIndex)
	End Property
	
	Private Sub InitDays()
		Dim nDayIndex
		Dim objNewDay
		
		If mcolDays.Count > 0 Then mcolDays.RemoveAll()
		
		For nDayIndex = 1 To mnDaysInMonth
			Set objNewDay = New CalendarDay
			objNewDay.DateString = FormatDateTime(DateSerial(mnYear, mnMonth, nDayIndex),2)
			objNewDay.OnClick = OnDayClick
			
			mcolDays.Add nDayIndex, objNewDay
		Next
		
		mbDaysInitialized = True
	End Sub
	
	Public Sub SetDate(dDate)
		mdDate  = CDate(dDate)
		mnDay   = Day(dDate)
		mnMonth = Month(dDate)
		mnYear  = Year(dDate)
	
		mnDaysInMonth =  Day(DateAdd("d", -1, DateSerial(mnYear, mnMonth + 1, 1)))
		mnDayMonthStarts = WeekDay(DateAdd("d", -(Day(CDate(dDate)) - 1), CDate(dDate)))
	End Sub
	
	Public Sub Draw()
		Dim nDayCount
		Dim nCellWidth, nCellHeight, nFontSizeRatio
		Dim objDay
		
		If Not mbDaysInitialized Then InitDays()
		
		nCellWidth = CInt(Width / 7)
		If (mnDayMonthStarts + mnDaysInMonth - 1) > 35 Then
			nCellHeight = CInt((Height - 80) / 6)
		Else
			nCellHeight = CInt((Height - 80) / 5)
		End If
		
		nFontSizeRatio = Fix(Width / 200)
		
		Send "<div id=""calendar"" style=""top: " & CStr(Top) & "px; left: " & CStr(Left) & "px; position: " & Position & "; z-index: " & ZIndex & """>"
		Send "<table border=""1"" width=""" & Width & """ height=""" & Height & """ cellspacing=""0"">"
		Send "<tr><td colspan=""7"" height=""10"" bgcolor=""" & TitlebarColor & """>"
		Send "	<table border=""0"" width=""100%"" cellspacing=0>"
		Send "	<tr>"
		Send "	<td align=""left""><a style=""text-decoration: none; color: " & TitlebarFontColor & ";"" href=""" & Replace(OnPrevMonthClick, "$date", DateSerial(mnYear, mnMonth - 1, mnDay)) & """><font face=""" & TitlebarFont & """ size=""" & nFontSizeRatio & """><b>&nbsp;&lt;&lt;</b></font></a></td>"
		Send "	<td align=""center""><font size=""" & nFontSizeRatio & """ face=""" & TitlebarFont & """ color=""" & TitlebarFontColor & """><b>" & MonthName(mnMonth) & " " & mnYear & "</b></font></td>"
		Send "	<td align=""right""><a style=""text-decoration: none; color: " & TitlebarFontColor & ";"" href=""" & Replace(OnNextMonthClick, "$date", DateSerial(mnYear, mnMonth + 1, mnDay)) & """><font face=""" & TitlebarFont & """ size=""" & nFontSizeRatio & """><b>&gt;&gt;&nbsp;</b></font></a></td>"
		Send "	</tr>"
		Send "	</table>"
		Send "</td></tr>"
		Send "<tr>"
		Send "<td height=""20"" width=""" & CStr(nCellWidth) & """ align=""center""><small>S</small></td>"
		Send "<td height=""20"" width=""" & CStr(nCellWidth) & """ align=""center""><small>M</small></td>"
		Send "<td height=""20"" width=""" & CStr(nCellWidth) & """ align=""center""><small>T</small></td>"
		Send "<td height=""20"" width=""" & CStr(nCellWidth) & """ align=""center""><small>W</small></td>"
		Send "<td height=""20"" width=""" & CStr(nCellWidth) & """ align=""center""><small>T</small></td>"
		Send "<td height=""20"" width=""" & CStr(nCellWidth) & """ align=""center""><small>F</small></td>"
		Send "<td height=""20"" width=""" & CStr(nCellWidth) & """ align=""center""><small>S</small></td>"
		Send "</tr>"
		
		Send "<tr>"
		For nDayCount = 1 To mnDayMonthStarts - 1
			Send "<td height=""" & CStr(nCellHeight) & """ width=""" & CStr(nCellWidth) & """ bgcolor=""#dddddd"">&nbsp;</td>"
		Next
		
		nDayCount = nDayCount - 1
		
		For Each objDay In mcolDays.Items
		
			If nDayCount = 7 Then 
				Send "</tr><tr>"
				nDayCount = 0
			End If	
			
			Response.Write "<td height=""" & CStr(nCellHeight) & """ width=""" & CStr(nCellWidth) & """ valign=""top"" bgcolor="""
			If objDay.DateString = msToday Then Send TodayBGColor & """>" Else Send "white"">"
			
			objDay.Draw()
			Send "</td>"
			
			nDayCount = nDayCount + 1
		Next

		If nDayCount < 7 Then
			For nDayCount = nDayCount To 6
				Send "<td height=""" & CStr(nCellHeight) & """ width=""" & CStr(nCellWidth) & """ bgcolor=""#dddddd"">&nbsp;</td>"
			Next
		End If
			
		Send "</tr>"
		
		If ShowDateSelect Then
			Send "<tr><td height=""30"" colspan=""7"" align=""center"">"
			DrawDateSelect()
			Send "</td></tr>"
		End If
		
		Send "</table>"
		Send "</div>"
	End Sub
	
	Private Sub DrawDateSelect()
		Dim nIndex
		Send "	<form id=frmGO name=frmGO>"
		Send "	<table border=""0"">"
		Send "	<tr>"
		Send "	<td><select name=""month"">"
			For nIndex = 1 To 12
				Response.Write "<option value=""" & nIndex & """" 
				If nIndex = Month(mdDate) Then Response.Write " selected"
				Send ">" & MonthName(nIndex, True) & "</option>"
			Next
		Send "	</select></td>"
		Send "	<td><select name=""year"">"
			For nIndex = Year(Now()) - 4 To Year(Now()) + 6
				Response.Write "<option value=""" & nIndex & """" 
				If nIndex = Year(mdDate) Then Response.Write " selected"
				Send ">" & CStr(nIndex) & "</option>"
			Next
		Send "	</select></td>"
		Send "	<td><input type=""button"" Value=""Go"" onclick=""document.location='" & Request.ServerVariables("SCRIPT_NAME") & "?date='+this.form.month.options[this.form.month.selectedIndex].value+'/1/'+this.form.year.options[this.form.year.selectedIndex].value;"" id=1 name=1></td>"
		Send "	</form>"
		Send "	</tr></table>"
	End Sub
	
	Private Sub Send(sHTML)
		Response.Write sHTML & vbCrLf
	End Sub

End Class


Class CalendarDay
	Public DateString
	Public OnClick
	Private mcolActivities
	Private mbActivitiesInit
	
	Private Sub Class_Initialize()
		mbActivitiesInit = False
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(mcolActivities) Then
			mcolActivities.RemoveAll()
			Set mcolActivities = Nothing
		End If
	End Sub
	
	Private Sub InitActivities()
		Set mcolActivities = Server.CreateObject("Scripting.Dictionary")
		mbActivitiesInit = True
	End Sub
	
	Public Sub AddActivity(sActivity, sColor)
		If Not mbActivitiesInit Then InitActivities()
		mcolActivities.Add mcolActivities.Count + 1, "bgcolor=""" & sColor & """>" & sActivity
	End Sub
	
	Public Sub Draw()
		Dim objActivity
		
		Send "<table width=""100%"" border=""0"" cellspacing=""2"" cellpadding=""1"">"
		Send "<tr><td align=""left"" valign=""top""><a href=""" & Replace(OnClick, "$date", DateString) & """><small>" & Day(DateString) & "</small></a></td></tr>"
		If mbActivitiesInit Then
			For Each objActivity In mcolActivities.Items
				Send "<tr><td height=""20""" & objActivity & "</td></tr>"
			Next
		End If
		Send "</table>"
	End Sub

	Private Sub Send(sHTML)
		Response.Write sHTML & vbCrLf
	End Sub
End Class


%>