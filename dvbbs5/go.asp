<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<%
	response.buffer=true
	dim times
	stats="��ת����"
	if request("boardid")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��ָ����̳���档"
	elseif not isInteger(request("boardid")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ��İ��������"
	else
		boardid=request("boardid")
	end if
	if request("sid")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
	elseif not isInteger(request("sid")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
	else
		times=request("sid")
	end if
	if founderr then
		call nav()
		call head_var(1)
		call dvbbs_error()
	else
		call nav()
		call head_var(2)
		call main()
		if founderr then call dvbbs_error()
	end if
	call footer()
	sub main()
	if request("action")="next" then
	set rs=conn.execute("select top 1 topicid from topic where boardid="&boardid&" and topicid>"&times&" and not locktopic=2 order by LastPostTime")
	if rs.eof and rs.bof then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>û���ҵ���һƪ���ӣ���<a href=list.asp?boardid="&boardid&">����</a>��"
	else
		response.redirect "dispbbs.asp?boardid="&boardid&"&ID="&rs(0)&""
	end if
	else
	set rs=conn.execute("select top 1 topicid from topic where boardid="&boardid&" and  topicid<"&times&" and not locktopic=2 order by LastPostTime desc")
	if rs.eof and rs.bof then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>û���ҵ���һƪ���ӣ���<a href=list.asp?boardid="&boardid&">����</a>��"
	else
		response.redirect "dispbbs.asp?boardid="&boardid&"&ID="&rs(0)&""
	end if
	end if
	end sub
	set rs=nothing
%>