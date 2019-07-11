<html>
<head>
<title>对字符串加密
</title>
</head>
<body>
<%
Const g_CryptThis = "下面对本句话进行加密了！"
Const g_KeyLocation = "c:\key.txt"
g_Key = mid(ReadKeyFromFile(g_KeyLocation),1,Len(g_CryptThis))
Response.Write "<p>原文: " & g_CryptThis & "<p>"
Response.Write "<p>密钥: " & g_Key & "<p>"
Response.Write "<p>加密: " & EnCrypt(g_CryptThis) & "<p>"
Response.Write "<p>解密: " & DeCrypt(EnCrypt(g_CryptThis)) & "<p>"
Function EnCrypt(strCryptThis)
Dim strChar, iKeyChar, iStringChar, I
for I = 1 to Len(strCryptThis)
iKeyChar = Asc(mid(g_Key,I,1))
iStringChar = Asc(mid(strCryptThis,I,1))
iCryptChar = iKeyChar Xor iStringChar
strEncrypted = strEncrypted & Chr(iCryptChar)
next
EnCrypt = strEncrypted
End Function
Function DeCrypt(strEncrypted)
Dim strChar, iKeyChar, iStringChar, I
for I = 1 to Len(strEncrypted)
iKeyChar = (Asc(mid(g_Key,I,1)))
iStringChar = Asc(mid(strEncrypted,I,1))
iDeCryptChar = iKeyChar Xor iStringChar
strDecrypted = strDecrypted & Chr(iDeCryptChar)
next
DeCrypt = strDecrypted
End Function
Function ReadKeyFromFile(strFileName)
Dim keyFile, fso, f
set fso = Server.CreateObject("Scripting.FileSystemObject")
set f = fso.GetFile(strFileName)
set ts = f.OpenAsTextStream(1, -2)
Do While not ts.AtEndOfStream
keyFile = keyFile & ts.ReadLine
Loop
ReadKeyFromFile = keyFile
End Function
%>
</body>
</html>