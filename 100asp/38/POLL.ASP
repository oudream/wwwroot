<html>
<title>
简单民意调查
</title>
<body>
<center>
<%
dim numberstring
dim numA,numB,numC,numD
whichfile=server.mappath("poll.txt")
set fs=createobject("Scripting.FileSystemObject")
set thisfile=fs.opentextfile(whichfile)
numberstring=thisfile.readline
thisfile.close
numA=clng(mid(numberstring,instr(numberstring,"A")+1,instr(numberstring,"B")-instr(numberstring,"A")-1))
numB=clng(mid(numberstring,instr(numberstring,"B")+1,instr(numberstring,"C")-instr(numberstring,"B")-1))
numC=clng(mid(numberstring,instr(numberstring,"C")+1,instr(numberstring,"D")-instr(numberstring,"C")-1))
numD=clng(mid(numberstring,instr(numberstring,"D")+1,len(numberstring)-instr(numberstring,"D")))

select case request.form("Editor")
    case "A": numA=numA+1
    case "B": numB=numB+1
    case "C": numC=numC+1
    case "D": numD=numD+1
end select
numberstring="A" & cstr(numA) & "B" & cstr(numB) & "C" & cstr(numC) & "D" & cstr(numD)
set out=fs.createtextfile(whichfile)
out.writeLine(numberstring)
out.close
set fs=nothing
response.write "<br>谢谢你的参与,"
response.write "调查结果为:"
%>
<table border=0>
<tr><td>JavaScript:</td><td><%=numA%></td></tr>
<tr><td>VBScript:</td><td><%=numB%></td></tr>
<tr><td>PerlScript:</td><td><%=numC%></td></tr>
<tr><td>其它:</td><td><%=numD%></td></tr>
</table>
</center>
</body>
</html>