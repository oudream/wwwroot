<html>
<head>
<title>AspChart组件示例(立体饼状图)</title>
</head>
<body bgcolor=white>
<blockquote>
<center>
<h2>AspChart组件示例(立体饼状图)</h2>
<%
  Dim Chart
  cBar    = 1
  cPie    = 3
  cNone    = 0
  cRectGradient =6
  Set Chart = Server.CreateObject ("ASPChart.Chart")
  Chart.AddSeries (cPie)
  Chart.AddValue 200, "第一小组", vbBlue
  Chart.AddValue 233, "第二小组", vbRed
  Chart.AddValue 220, "第三小组", vbGreen
  Chart.PanelColor = vbWhite
  Chart.BevelOuter = cNone
  Chart.ChartBGColor = vbWhite
  Chart.HintsVisible = true
  Chart.Height = 200
  Chart.Width = 550
  Chart.FileName = "c:\inetpub\wwwroot\images\smpchrt5.jpg"
  Chart.SaveChart
  Response.Write "<img src=""/images/smpchrt5.jpg""><p><br></p><p><br></p>"
  Set Chart = nothing
%>
</center>
</blockquote>
</body>
</html>