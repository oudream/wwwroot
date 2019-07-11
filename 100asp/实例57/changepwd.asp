<html>
<title>
更改NT用户密码
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
Response.Write "<font color=red><b>密码已经成功修改! </b><br><br>"
Response.Write "</font>"
Else
objContext.SetAbort
Response.Write "<font color=red><b>错误的旧密码，请重新输入正确的旧密码!</b><br><br>"&ErrMsg
Response.Write "</font>"
End If
Set adsUser=Nothing
End Sub 

response.write "<center><h2>更改NT用户密码</h2></center>"
Computer="fang"
UserName="fenfang"
oldPassword="1234"
newPassword="4321"
ChangeUserPassword Computer,UserName,oldPassword,newPassword
%>
</body>
</html>
