<!--#include file=conn.asp-->
<!--#include file=inc/const.asp-->
<html>
<body>
<%Response.Expires=0%>
<%
dim outtext
dim num		'倒数第几个跟帖
dim allnum	'总共多少跟帖
dim star		'第几页
dim rootID

if request("rootID")="" then
	response.write "非法的贴子参数。"
	response.end
elseif not isInteger(request("rootID")) then
	response.write "非法的贴子参数。"
	response.end
else
	rootID=request("rootID")
end if
num=0
outtext="&nbsp;&nbsp;"
function HTMLEncode(fString)
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = replace(fString, "'", "")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "；")
    fString = Replace(fString, CHR(10), "，")
    HTMLEncode = fString
end function
dim totalusetable

set rs=conn.execute("select child,PostTable from topic where topicid="&rootid)
allnum=rs(0)
totalusetable=rs(1)

sql="select layer,rootid,announceid,topic,body,username,postuserid from "&totalusetable&" where boardid="& boardid &" and rootid="& rootid &" and parentid>0 and not locktopic=2 order by rootid desc,orders"
set rs=conn.execute(sql)
%>
<script>
	parent.followTd<%=rootid%>.innerHTML='<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%" align=center><TBODY><%do while not rs.eof%><TR><TD class=tablebody1 width="100%" height=25><%for i=2 to rs("layer")%><%=outtext%><%next%><img src="<%=Forum_info(7)%>nofollow.gif"><a href="dispbbs.asp?boardID=<%=boardID%>&ID=<%=rs("RootID")%>&replyID=<%=rs("announceID")%><%
	star=int((allnum-num)/Forum_Setting(12))+1
	if  star>1 then
		response.write "&star="&star
	end if%>&skin=0#<%=rs("announceid")%>" title="<%=htmlencode(rs("topic"))%>"><%
	if rs("topic")="" then
	response.write left(htmlencode(rs("body")),22)&"..."
	else
	if len(rs("topic"))>22 then
		response.write ""&htmlencode(left(rs("topic"),22))&"..."
	else
		response.write htmlencode(rs("topic"))
	end if
end if
%></font></a> -- <a href="dispuser.asp?id=<%=htmlencode(rs("postuserid"))%>"><%=htmlencode(rs("username"))%></a></tr><%
	
	rs.movenext
	num=num+1
	loop
	rs.close
	conn.close
	set conn=nothing%></TBODY></TABLE>';
</script>
</body>
</html>