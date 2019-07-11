<%@ Language="VBScript" %>
<% Option Explicit %>
<html>
<head>
    <title>²é¿´ÍøÂçÉèÖÃ</title>
</head>
<body bgcolor="#FFFFFF">
<%
    dim strHost
    dim oShell,oFS,oTF
    dim i,Data,tempData
    strHost="ipconfig"
    Set oShell = Server.CreateObject("Wscript.Shell")
    oShell.Run "%ComSpec% /c ipconfig > C:\" & strHost & ".txt", 0, True
    Set oFS = Server.CreateObject("Scripting.FileSystemObject")
    Set oTF = oFS.OpenTextFile("C:\" & strHost & ".txt")
    Do While Not oTF.AtEndOfStream
        Data = Trim(oTF.Readline)
            If i > 2 Then 
                tempData = tempData & Data & "<BR>"
            End If
        i = (i + 1)
    Loop
    response.write tempData
    oTF.Close
    oFS.DeleteFile "C:\" & strHost & ".txt"
    Set oFS = Nothing
%>
</font>
</body>
</html>

 
