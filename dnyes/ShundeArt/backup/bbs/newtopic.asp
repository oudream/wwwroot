<!--#include file="conn.asp"-->
<!--#Include File="inc/Dv_ClsMain.asp"-->
<%
'DVBBS 7.0 ������̳��ҳ����-----�������
dim bbsurl,lockboardid,picurl
'############����Ϊ�޸���######################
dim lockurl
lockurl=""
'ֻ���������ַ,Ҫ��"HTTP://"��ͷ,Ϊ���򲻿��Ŵ˹���.(���������ַ���ƣ�Ҫ��","�ָ���)
'����ֻ�����������ַ����: lockurl="http://www.artistsky.net/,http://www.artbbs.net/"
bbsurl="http://joy.js.zwu.edu.cn/ft-bbs/"							'����д����̳����ȷ��ַ,Ҫ��"HTTP://"��ͷ
lockboardid=""										'����д���Ƶ��õ���̳���ID,�ö��Ÿ���������lock����Ϊ1��2ʱ��Ч��
picurl="http://joy.js.zwu.edu.cn/ft-bbs/skins/default/topicface/"	'����ͼ��Ŀ¼��ַ
'############����Ϊ�޸���######################
'bbsurl=getservepath(request.ServerVariables("server_name")&request.ServerVariables("URL"))
'function getservepath(str)
'dim tmpstr
'tmpstr=split(str,"/")
'getservepath="http://"&replace(str, tmpstr(ubound(tmpstr)), "")
'end function
'*************************************
'�ϴ�����CONN.ASPͬ����Ŀ¼��
'���ϵ�ַ����һ��Ҫ�޸�,���������õ�������ȥ�����ϵ���̳.
'��������,��������һ���ϴ���newscode.ASP�ļ����е���(newscode.ASP����ǰҪ�޸ĵ��ò���)
'	FSSUNWIN	2003.12.31
'*************************************
if trim(lockurl)<>"" and checkserver(lockurl)=false then
	response.write "document.write ('���ݱ�����,��ֹ������վ�����!');"
	response.end	
end if

Private function checkserver(str)
	dim i,servername
	checkserver=false
	if str="" then exit function
	str=split(Cstr(str),",")
	servername=Request.ServerVariables("HTTP_REFERER")
	for i=0 to Ubound(str)
	if right(str(i),1)="/" then str(i)=left(trim(str(i)),len(str(i))-1)
		if Lcase(left(servername,len(str(i))))=Lcase(str(i)) then
			checkserver=true
			exit for
		else
			checkserver=false
		end if
	next
end function

dim rs,sql
dim orders,reply,topic,isbest,lock,board
dim bname,ars
dim postinfo,postname,POSTTIME
dim NowUseBbs,boardname,boardid
dim i,k,n,sdate,searchdate
	i=0:k=0
	lock=cint(trim(request("lock")))

	if trim(request("n"))<>"" and IsNumeric(request("n")) then
	n=cint(request("n"))
	else
	n=1
	end if

	if trim(request("orders"))=1 then
		orders="hits desc,"
	Elseif trim(request("orders"))=2 or trim(request("orders"))=3 then
		orders="dateandtime desc,"
	end if
	boardid=trim(request("boardid"))
	If boardid<>"all" and isnumeric(boardid) then
		if boardid=444 then
		response.write "document.write ('����İ����������ñ���ֹ��');"
		response.end
		Else
		board=" and BoardID="&cint(boardid)
		if lock=3 then board=" and BoardID in (select boardid from board where ParentID="&cint(boardid)&") "
		End If
	End If
	
	if lock=1 then
		board=" and boardid not in ("&lockboardid&") "
	elseif lock=2 then
		board=" and boardid in ("&lockboardid&") "
	end if

	Dvbbs.GetForum_Setting
	connectionDatabase
	if trim(request("sdate"))<>"" and IsNumeric(request("sdate")) then
		sdate=cint(request("sdate"))
		if IsSqlDataBase=1 Then
		searchdate=" and datediff(day,dateandtime,"&SqlNowString&")<"&sdate
		else
		searchdate=" and datediff('d',dateandtime,"&SqlNowString&")<"&sdate
		end if
	else
	searchdate=""
	end if

	if cint(request("action"))=1 then
		'��ʾ����
		if trim(request("orders"))=2 then orders="lastposttime"
		set rs=conn.execute("select top "&n&" PostUserName,Title,topicid,boardid,dateandtime,topicid,hits,Expression,LastPost from Dv_topic where boardid<>444 "&board&searchdate&" ORDER BY "&orders&" topicid desc")
	elseif cint(request("action"))=2 then
		'��ʾ��������
		if searchdate<>"" then searchdate=replace(searchdate," and"," where")
		if searchdate="" and board<>"" then board=replace(board," and"," where")
		set rs=conn.execute("select top "&n&" PostUserName,Title,rootid,boardid,dateandtime,Announceid,id,Expression from Dv_BestTopic  "&board&searchdate&"  ORDER BY "&orders&" id desc")
    else
	    '��ʾ�����ظ�
		set rs=conn.execute("select top "&n&" username,topic,rootid,boardid,dateandtime,announceid,body,Expression from "&Dvbbs.NowUseBBS&" where (not boardid=444) "&board&searchdate&" ORDER BY "&orders&" AnnounceID desc")
	end if
	If Not RS.Eof then
		SQL=Rs.GetRows(-1)
		else
		response.write "document.write('��δ�������ӣ�');"
		response.end
    end if
    rs.close
	set rs=nothing

	For i=0 To Ubound(SQL,2)
		topic=SQL(1,i)
		if topic="" then
			topic=SQL(6,i)
		end if
		Topic=Stringhtml(topic)
		if len(topic)>Cint(request("tlen")) then
			topic=left(topic,request("tlen"))&"..."
		end if
		
		postname=SQL(0,i)
		POSTTIME=SQL(4,i)
		if request("action")=1 and request("reply")=1 then
			if SQL(8,i)<>"" then
			postinfo=split(SQL(8,i),"$")
			postname=postinfo(0)
			POSTTIME=postinfo(2)
			end if
		end if
		if request("showpic")=1 then
		response.write "document.write('<IMG SRC="""&picurl&SQL(7,i)&""" BORDER=0 > ');"
		else
		response.write "document.write('<font color=#000000>��</FONT> ');"
		end if
		if request("bname")=1 then
			set ars=conn.execute("select BoardType from Dv_board where boardid="&SQL(3,i))
				boardname=ars(0)
				ars.close
				response.write "document.write('<a href="&bbsurl&"list.asp?boardid="&SQL(3,i)&" target=""_blank"">"&Dvbbs.htmlencode(boardname)&"</a>');"
		end if
		response.write "document.write('<a href="&bbsurl&"dispbbs.asp?boardid="&SQL(3,i)&"&ID="&SQL(2,i)&"&replyID="&SQL(5,i)&" target=""_blank"" title="&Topic&" class=middle>');"
		response.write "document.write('"&Topic&"');"
		response.write "document.write('</a>');"
		select case cint(request("info"))
		case 0
		case 1
		response.write "document.write('��(<a href="&bbsurl&"dispuser.asp?name="&postname&" target=_blank>"&postname&"</a>,<font color=green>"&formatdatetime(POSTTIME,0)&"</font>)');"
		case 2
		response.write "document.write('��(<font color=green>"&POSTTIME&"</font>)');"
		case 3
		response.write "document.write('��(<a href="&bbsurl&"dispuser.asp?name="&postname&" target=_blank>"&postname&"</a>)');"
		case 4
		response.write "document.write('��(<a href="&bbsurl&"dispuser.asp?name="&postname&" target=_blank>"&postname&"</a>');"
		if cint(request("action"))=1 then response.write "document.write(',<font color=green>"&SQL(6,i)&"</font>');"
		Response.Write "document.write(')');"
		case 5
		if cint(request("action"))=1 then
		response.write "document.write('(<font color=green>"&SQL(6,i)&"</font>)');"
		end if
		case 6
		response.write "document.write('��(<a href="&bbsurl&"dispuser.asp?name="&postname&" target=_blank>"&postname&"</a>,<font color=green>"&formatdatetime(POSTTIME,1)&"</font>)');"
		case 7
		response.write "document.write('��(<font color=green>"&formatdatetime(POSTTIME,1)&"</font>)');"
		case else

		end select
		response.write "document.write('<br>');"
		k=k+1
	Next
	Call CloseObject

	Sub CloseObject()
		Set template = Nothing
		Set MyBoardOnline = Nothing
		Set Dvbbs = Nothing
		Set Conn = Nothing
	End Sub

	Function Stringhtml(str)
		Dim re
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		're.Pattern="<(.*)>"
		'str=re.replace(str, "")
		re.Pattern="\[(.*)\]"
		str=re.replace(str, "")
		str = Replace(str, CHR(34), """")
		str = Replace(str, CHR(39), "\'")
		str = Replace(str, CHR(13), "")
		str = Replace(str, CHR(10), "")
		str = replace(str, ">", "&gt;")
		str = replace(str, "<", "&lt;")
		if str="" then str="..."
		Stringhtml=str
	End Function
%>

