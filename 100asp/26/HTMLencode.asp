<Html>
<Head>
<Title>
��ʾ������
</Title>
</Head>
<Body>
<%
    s_message="<a href=hello.asp>hello</a>"
    Response.Write "����ת��ǰ�����:<br>"
    Response.Write s_message & "<br>"
    Response.Write "������to_html����ת��������:<br>"  
    Response.Write to_html(s_message) & "<br>"
    Response.Write "������Server.HTMLencode����ת��������:<br>"  
    Response.Write Server.HTMLEncode(s_message)
%>


<%
Function to_html(s_string)
    to_html = Replace(s_string, """", "&quot;")
    to_html = Replace(to_html, "<", "&lt;")
    to_html = Replace(to_html, ">", "&gt;")
    to_html = Replace(to_html, vbcrlf, "<br>")
    to_html = Replace(to_html, "/&lt;", "<")
    to_html = Replace(to_html, "/&gt;", ">")
    to_html = edit_hrefs(to_html)
End Function
%>

<script language="javascript1.2" runat=server>
function edit_hrefs(s_html){
    // һ��ʹ��������ʽ�ĵ䷶
    // ת���ı������еĳ����Ӻ͵����ʼ���ʽ
    s_str = new String(s_html);

    s_str = s_str.replace(/\bhttp\:\/\/www(\.[\w+\.\:\/\_]+)/gi, 
        "http\:\/\/&not;��&cedil;$1");

    s_str = s_str.replace(/\b(http\:\/\/\w+\.[\w+\.\:\/\_]+)/gi,
        "<a href=\"$1\">$1<\/a>");
        
    s_str = s_str.replace(/\b(www\.[\w+\.\:\/\_]+)/gi, 
        "<a href=\"http://$1\">$1</a>");
        
    s_str = s_str.replace(/\bhttp\:\/\/&not;��&cedil;(\.[\w+\.\:\/\_]+)/gi,
        "<a href=\"http\:\/\/www$1\">http\:\/\/www$1</a>");
        
    s_str = s_str.replace(/\b(\w+@[\w+\.?]*)/gi, 
        "<a href=\"mailto\:$1\">$1</a>");
        
    
    return s_str;
}
</script> 
</Body>
</Html>