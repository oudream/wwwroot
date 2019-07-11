<html>
<head>
<title>AspChart组件示例(条形图)</title>
</head>
<body>
<blockquote>
<center>
<h2>AspChart组件示例(条形图)</h2>
<%
  Dim Chart
  cNone  = 0
  cBar   = 1
  cRectGradient =6
  Set Chart = Server.CreateObject ("ASPChart.Chart")
  Chart.AddSeries (cBar)
  Chart.AddValue 200, "标准", vbBlue
  Chart.AddValue 233, "增强 ", vbRed
  Chart.AddValue 260, "免费", vbGreen
  Chart.BarStyle = cRectGradient
  Chart.ChartTitleAdd ("简单条形图")
  Chart.ChartTitleFont.Size = 20
  Chart.ChartTitleFont.Name = "Times New Roman"
  Chart.ChartTitleFont.Italic = true
  Chart.ChartTitleFont.Bold = true
  Chart.BevelOuter = cNone
  Chart.GradientVisible = true
  Chart.GradientStartColor = vbWhite
  Chart.GradientEndColor = vbYellow
  Chart.Height = 300
  Chart.Width = 500
  Chart.FileName = "c:\inetpub\wwwroot\images\smpchrt1.jpg"
  Chart.SaveChart
  Response.Write "<img src=""/images/smpchrt1.jpg"">"
  Set Chart = nothing
%>
</center>
</blockquote>
</body>
</html>