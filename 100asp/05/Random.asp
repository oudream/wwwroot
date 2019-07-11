<%@ LANGUAGE = VBScript %>
<html>
<title>
生成随机字符串
</title>
<body>
<%
Function gen_key(digits)
'定义并初始化数组
    dim char_array(80)
	'初始化数字
    For i = 0 To 9
        char_array(i) = CStr(i)
    Next
	'初始化大写字母
    For i = 10 To 35
        char_array(i) = Chr(i + 55)
    Next
	'初始化小写字母
    For i = 36 To 61
        char_array(i) = Chr(i + 61)
    Next
	Randomize   '初始化随机数生成器。
	do while len(output) < digits
        num = char_array(Int((62 - 0 + 1) * Rnd + 0))
        output = output + num
    loop
'设置返回值
    gen_key    =    output
End Function
'把结果返回给浏览器
response.write "本实例生成的十三位随机字符串为:"
response.write "<center>"
response.write gen_key(13)
response.write "</center>"
%>
</body>
</html>
