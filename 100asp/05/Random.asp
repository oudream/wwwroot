<%@ LANGUAGE = VBScript %>
<html>
<title>
��������ַ���
</title>
<body>
<%
Function gen_key(digits)
'���岢��ʼ������
    dim char_array(80)
	'��ʼ������
    For i = 0 To 9
        char_array(i) = CStr(i)
    Next
	'��ʼ����д��ĸ
    For i = 10 To 35
        char_array(i) = Chr(i + 55)
    Next
	'��ʼ��Сд��ĸ
    For i = 36 To 61
        char_array(i) = Chr(i + 61)
    Next
	Randomize   '��ʼ���������������
	do while len(output) < digits
        num = char_array(Int((62 - 0 + 1) * Rnd + 0))
        output = output + num
    loop
'���÷���ֵ
    gen_key    =    output
End Function
'�ѽ�����ظ������
response.write "��ʵ�����ɵ�ʮ��λ����ַ���Ϊ:"
response.write "<center>"
response.write gen_key(13)
response.write "</center>"
%>
</body>
</html>
