<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="calendar.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>面向对象日历</TITLE>
</HEAD>
<BODY LINK="blue" ALINK="blue" VLINK="blue">
<%
Dim MyCalendar
Set MyCalendar = New Calendar
MyCalendar.Top = 50
MyCalendar.Left = 150
MyCalendar.Position = "absolute" 
MyCalendar.Height = "200" 
MyCalendar.Width = "300" 
MyCalendar.TitlebarColor = "darkblue" 
MyCalendar.TitlebarFont = "arial" 
MyCalendar.TitlebarFontColor = "white" 
MyCalendar.TodayBGColor = "skyblue" 
MyCalendar.ShowDateSelect = True 
MyCalendar.OnDayClick = "javascript:alert('你点击了: $date')"
Select Case Month(MyCalendar.GetDate())
  Case 1
    MyCalendar.Days(1).AddActivity "<small><b>New Years</b></small>", "limegreen"
  Case 12
    MyCalendar.Days(25).AddActivity "<small><b>Christmas</b></small>", "limegreen"
End Select
MyCalendar.Draw()
%>
</BODY>
</HTML>
