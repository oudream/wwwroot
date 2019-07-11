<%@ Language=VBScript %>
<%Response.Buffer=true
%>

<HTML>
<HEAD>
<title>ÁÄÌìÊÒ</title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>

<frameset rows="80%,*">
	<frameset cols="75%,*">
		<frame name="left" target="lefttop"scrolling="auto" noresize src="Content.asp">
		<frame name="right" noresize scrolling="auto" noresize src="message.asp">
	</frameset>
		<frame name="bottom" target="righttop" scrolling="auto" noresize src="send.asp">
	<noframes>
	<body>
	<p>This page user frames,but your browser donesn't support them!
	</body>
	<noframes>
</frameset>	
</html>
