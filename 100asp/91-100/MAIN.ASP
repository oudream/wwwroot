<html>
<head>
<title>
</title>
</head>
<frameset frameborder=0 border=0 framespacing=0 frameborder=0 rows="100,*">
<%if request.querystring("register")=0 then%>
    <frame src="top.asp?register=0" name="menu" scrolling="no" marginwidth=0 magingeight=0>
<%else%>
    <frame src="top.asp?register=1" name="menu" scrolling="no" marginwidth=0 magingeight=0>   
<%end if%>
<frame src=bottom.htm name=main>
<noframes>
<p>This page uses frames,but your bowser doesn't support them.</p>
</body>
</noframes>
</frameset>
</html>