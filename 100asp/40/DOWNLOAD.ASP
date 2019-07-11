<%@language=VBscript%>
<%
Const ForReading=1
Const TristateTrue=-1 
Const FILE_TRANSFER_SIZE=16384 
Response.Buffer = True
Function TransferFile(path, mimeType, filename)
Dim objFileSystem, objFile, objStream
Dim char
Dim sent
send=0
TransferFile = True
Set objFileSystem = Server.CreateObject("Scripting.FileSystemObject")
Set objFile = objFileSystem.GetFile(Path)
Set objStream = objFile.OpenAsTextStream(ForReading, TristateTrue)
Response.AddHeader "content-type", mimeType
response.AddHeader "Content-Disposition","attachment;filename=" & filename 
Response.AddHeader "content-length", objFile.Size
Do While Not objStream.AtEndOfStream
    char = objStream.Read(1)
    Response.BinaryWrite(char)
    sent = sent + 1
    If (sent MOD FILE_TRANSFER_SIZE) = 0 Then
        Response.Flush
        If Not Response.IsClientConnected Then
            TransferFile = False
            Exit Do
        End If
    End If
Loop
Response.Flush
If Not Response.IsClientConnected Then TransferFile = False
objStream.Close
Set objStream = Nothing
Set objFileSystem = Nothing
End Function
Dim path, mimeType, sucess,downfilename
downfilename=request("filename")
path = Server.MapPath(downfilename)
mimeType="text/plain"
sucess = TransferFile(path, mimeType,downfilename)
Response.End
%>
