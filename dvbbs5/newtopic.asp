<!--#include file="conn.asp"-->
<%
function HTMLEncode(fString)

    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = replace(fString, ";", "£»")
    fString = replace(fString, "'", "¡®")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "£»")
    fString = Replace(fString, CHR(10), "£¬")
    HTMLEncode = fString
end function
	dim rs,sql
	set rs=server.createobject("adodb.recordset")
	'sql="SELECT top "&request("n")&" username,topic,announceid,boardid,rootid,body FROM bbs1 where parentid=0 ORDER BY times desc,hits DESC"
	sql="select top "&request("n")&" username,topic,announceid,boardid,rootid,hits from bbs1 where parentID=0 and datediff('d',dateandtime,Now())=0 and not locktopic=2 ORDER BY hits desc"
	rs.open sql,conn,1,1
	do while Not RS.Eof
	response.write "document.write('<FONT color=#b70000><B>¡¤</B></FONT><span style=""font-size:9pt;line-height: 15pt""><a href=http://bbs.aspsky.net/dispbbs.asp?boardid="&rs(3)&"&RootID="&rs(4)&"&ID="&rs(2)&" target=_top title="&htmlencode(rs(1))&">');"

	if len((rs("topic")))>16 then
		response.write "document.write('"&htmlencode(left(rs(1),16))&"...');"
	else
		response.write "document.write('"&htmlencode(rs(1))&"');"
	end if
	response.write "document.write('</a>£ (<a href=dispuser.asp?name="&rs(0)&" target=_blank>"&rs(0)&"</a>,<font color=green>"&rs(5)&"</font>)</span><br>');"
	RS.MoveNext
	Loop
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>

