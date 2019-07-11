<html>
<head>
<title>
图像计数器
</title>
</head>
<body>
<br><br><br><br><br><br><br><br><br><br><hr>
<%
dim visitors
whichfile=server.mappath("counter.txt")

set fs=createobject("Scripting.FileSystemObject")
set thisfile=fs.opentextfile(whichfile)
visitors=thisfile.readline
thisfile.close

CountLen=len(visitors)
response.write "<center>你是第"
for i=1 to 6-countLen
	response.write "<img src=num\0.gif></img>"
next 

for i=1 to countlen
	response.write "<img src=num\" & mid(visitors,i,1) & ".gif></img>"
next 
response.write "位访问本站</center>"
visitors=visitors+1

set out=fs.createtextfile(whichfile)
out.writeLine(visitors)
out.close

set fs=nothing
%>
</body>
</html>