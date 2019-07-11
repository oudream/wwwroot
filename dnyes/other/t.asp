<%
'常用函数

'1、输入url目标网页地址，返回值getHTTPPage是目标网页的html代码
function getHTTPPage(url)
    dim Http
    set Http=server.createobject("MSXML2.XMLHTTP")
    Http.open "GET",url,false
    Http.send()
    if Http.readystate<>4 then 
        exit function
    end if
    getHTTPPage=bytesToBSTR(Http.responseBody,"GB2312")
    set http=nothing
    if err.number<>0 then err.Clear 
end function

'2、转换乱玛，直接用xmlhttp调用有中文字符的网页得到的将是乱玛，可以通过adodb.stream组件进行转换
Function BytesToBstr(body,Cset)
        dim objstream
        set objstream = Server.CreateObject("adodb.stream")
        objstream.Type = 1
        objstream.Mode =3
        objstream.Open
        objstream.Write body
        objstream.Position = 0
        objstream.Type = 2
        objstream.Charset = Cset
        BytesToBstr = objstream.ReadText 
        objstream.Close
        set objstream = nothing
End Function




'下面试着调用http://www.3doing.com/earticle/的html内容
Dim Url,Html
Url="http://www.3doing.com/earticle/"
Html = getHTTPPage(Url)
Response.write Html
%>