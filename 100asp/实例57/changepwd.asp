<html>
<title>
����NT�û�����
</title>
<body>
<%
Sub ChangeUserPassword(Computer,UserName,oldPassword,newPassword)
Dim adsUser,foundErr,ErrMsg
On Error Resume Next
foundErr=False
ErrMsg=""
Set adsUser=GetObject("WinNT://"+Computer+"/"+UserName+",user")
If Err.Number<>0 Then
foundErr=True
ErrMsg="User not found!"
Err.Clear
Else
adsUser.ChangePassword oldPassword, newPassword
adsUser.SetInfo
If Err.Number<>0 Then
foundErr=True
ErrMsg=Now & ". Error Code: " & Hex(Err) & " - " & Err.Description & "<br>"
Err.Clear
End If
End If
If Not foundErr Then
objContext.SetComplete
Response.Write "<font color=red><b>�����Ѿ��ɹ��޸�! </b><br><br>"
Response.Write "</font>"
Else
objContext.SetAbort
Response.Write "<font color=red><b>����ľ����룬������������ȷ�ľ�����!</b><br><br>"&ErrMsg
Response.Write "</font>"
End If
Set adsUser=Nothing
End Sub 

response.write "<center><h2>����NT�û�����</h2></center>"
Computer="fang"
UserName="fenfang"
oldPassword="1234"
newPassword="4321"
ChangeUserPassword Computer,UserName,oldPassword,newPassword
%>
</body>
</html>
