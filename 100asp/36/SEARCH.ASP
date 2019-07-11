<HTML>
<HEAD>
<TITLE>'<%=Request("SearchText")%>'的搜索结果</TITLE>
</HEAD>
<BODY>
<B>'<%=Request("SearchText")%>'的搜索结果</B><BR>
<%
Const fsoForReading = 1
Dim objFile, objFolder, objSubFolder, objTextStream
Dim bolCase, bolFileFound, bolTagFound
Dim strCount, strDeTag, strExt, strFile, strContent, strRoot, strTag, strText, strTitle, strTitleL
strFile = ".asp .htm .html .js .txt .css"
strRoot = "/"
strText = Request("SearchText")
strTag = Chr(37) & Chr(62)
bolFileFound = False
bolTagFound = False
If Request("Case") = "on" Then bolCase = 0 Else bolCase = 1
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
RealPath=Server.MapPath(strRoot)
VirtualPath="http://" & Request.ServerVariables("SERVER_NAME") 
Set objFolder = objFSO.GetFolder(RealPath)
schSubFol(objFolder)
Sub schSubFol(objFolder)
on error resume next
For Each objFile in objFolder.Files
  If strText = "" Then Exit Sub
  If Response.IsClientConnected Then
    Set objTextStream = objFSO.OpenTextFile(objFile.Path,fsoForReading)
    strContent = objTextStream.ReadAll
    If InStr(1, strContent, strTag, bolCase) Then
    Else
      If Mid(objFile.Name, Len(objFile.Name) - 1, 1) = "." Then strExt = Mid(objFile.Name, Len(objFile.Name) - 1, 2)
      If Mid(objFile.Name, Len(objFile.Name) - 2, 1) = "." Then strExt = Mid(objFile.Name, Len(objFile.Name) - 2, 3)
      If Mid(objFile.Name, Len(objFile.Name) - 3, 1) = "." Then strExt = Mid(objFile.Name, Len(objFile.Name) - 3, 4)
      If Mid(objFile.Name, Len(objFile.Name) - 4, 1) = "." Then strExt = Mid(objFile.Name, Len(objFile.Name) - 4, 5)
      If InStr(1, strContent, strText, bolCase) And Instr(1, strFile, strExt, 1) Then
        If InStr(1, strContent, "<TITLE>", 1) Then 
          strTitle = Mid(strContent, InStr(1, strContent, "<TITLE>", 1) + 7, InStr(1, strContent, "</TITLE>", 1)) 
        Else 
          strTitle = "未命名"
        end if
        myFile=objFile.Path
        myFile=replace(myfile,RealPath,VirtualPath,1,-1,1)
        myFile=replace(myfile,"\","/")
        strCount = strCount + 1
        Response.Write "<DL><DT><B><I>"& strCount  &"</I></B> - <A HREF=" & myFile & ">" & strTitle & "</A></A></DT><BR><DD>" 
        strTitleL = InStr(1, strContent, "</TITLE>", 1) - InStr(1, strContent, "<TITLE>", 1) + 7
        strDeTag = ""
        bolTagFound = False
        Do While InStr(strContent, "<")
          bolTagFound = True
          strDeTag = strDeTag & " " & Left(strContent, InStr(strContent, "<") - 1)
          strContent = MID(strContent, InStr(strContent, ">") + 1)
        Loop
        strDeTag = strDeTag & strContent
        If Not bolTagFound Then strDeTag = strContent
        Response.Write replace(Mid(strDeTag, strTitleL, 200),strText,"<font color=red>" & strText & "</font>",1,-1,bolcase) 
        Response.Write "...<BR><b><FONT SIZE='2'>URL: " & myFile 
        Response.Write " - 上次修改时间: " & objFile.DateLastModified
        Response.Write " - " & FormatNumber(objFile.Size / 1024)
        Response.Write "Kbytes</FONT></b></DD></DL>" 
        bolFileFound = True
      End If
      objTextStream.Close
    End If
  End If
Next
End Sub
For Each objSubFolder in objFolder.SubFolders
    schSubFol(objSubFolder)
Next
If Not bolFileFound then Response.Write "没有匹配结果"
If bolFileFound then Response.Write "<B>搜索结束</B>"
Set objTextStream = Nothing
Set objFolder = Nothing
Set objFSO = Nothing
%>
</BODY></HTML> 
