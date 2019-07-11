<HTML>  
<BODY>  
<%  

Set Upload = Server.CreateObject("Persits.Upload.1")  
Upload.Save "c:\upload"
%>  
Files:<BR> 
<%  

For Each File in Upload.Files  
Response.Write File.Name & "= " & File.Path & " (" & File.Size &")<BR>"
Next
%>  
<P>  
Otheritems:<BR>  
<%  
For Each Item in Upload.Form  
Response.Write Item.Name & "= " & Item.Value &"<BR>"
Next
%>  
</BODY>  
</HTML>
