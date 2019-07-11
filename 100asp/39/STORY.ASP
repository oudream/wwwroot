<%
If not request.Form("NextLine")="" then
  Set mFileObject=Server.CreateObject("Scripting.FileSystemobject")
  set mFile=mFileObject.OpenTextFile("d:\Story.txt",8,Ture)
  mFile.WriteLine(Request.Form("NextLine"))
  mFile.Close
end if
%>
<Html>
<Head>
<Title>
在线故事接龙
</Title>
</Head>
<Body>
<Center><h2>在线故事接龙</h2></Center>
<hr>
<%
Set mFileObject=Server.CreateObject("Scripting.FileSystemObject")
set mFile=mFileObject.OpenTextFile("D:\story.txt")
while not mFile.AtEndOfStream
  Response.Write "&nbsp;&nbsp;" & mFile.ReadLine
wend
mFile.close
%>
<hr>
<h3>请输入这个故事的新行:</h3>
<form method=post action=story.asp>
<input name="NextLine" type=text size=70>
<input type=submit value=确认添加本句>
</form>
</html>


