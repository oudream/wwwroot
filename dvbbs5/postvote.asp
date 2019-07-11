<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<%
	response.buffer=true
	dim voteid
	dim announceid
	dim action
	dim vote,votenum
	dim postvote(50)
	dim postvote1
	dim j,votenum_1,votenumlen
	dim vrs
	dim postnum,postoption
	stats="投票"
	if not founduser then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请登陆后投票。"
	end if
	if BoardID="" or not isInteger(BoardID) or BoardID=0 then
		Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
		founderr=true
	else
		BoardID=clng(BoardID)
	end if
	if request("id")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请指定相关贴子。"
	elseif not isInteger(request("id")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>非法的贴子参数。"
	else
		AnnounceID=request("id")
	end if
	if request("voteid")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请指定相关贴子。"
	elseif not isInteger(request("voteid")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>非法的贴子参数。"
	else
		voteID=request("voteid")
	end if
	if Cint(GroupSetting(9))=0 then
		Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛发表投票的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
		founderr=true
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
call activeonline()
call footer()

sub main()
	set rs=conn.execute("select locktopic from topic where topicid="&AnnounceID)
	if not (rs.eof and rs.bof) then
		if rs(0)=1 then
		Errmsg=Errmsg+"<br>"+"<li>您选择的投票已经关闭。"
		Founderr=true
		exit sub
		end if
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from vote where voteid="&voteid
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br>"+"<li>请您选择投票的主题进行投票。"
		Founderr=true
		exit sub
	else
		set vrs=conn.execute("select userid from voteuser where voteid="&voteID&" and userid="&userid)
		if not(vrs.eof and vrs.bof) then
			Errmsg=Errmsg+"<br>"+"<li>您已经投过票了。"
			Founderr=true
			exit sub
		else
			votenum=split(rs("votenum"),"|")
			if rs("votetype")=1 then
				for i = 0 to ubound(votenum)
				postvote(i)=request("postvote_"&i&"")
				next
			end if
			for j = 0 to ubound(votenum)
			if rs("votetype")=0 then
				if cint(request("postvote"))=j then
				votenum(j)=votenum(j)+1
				postoption=j
				end if
				votenum_1=""&votenum_1&""&votenum(j)&"|"
				postnum=1
			else
				if postvote(j)<>"" then
					if cint(postvote(j))=j then
						votenum(j)=votenum(j)+1
						postnum=postnum+1
						postoption=postoption & j & ","
					end if
				end if
				votenum_1=""&votenum_1&""&votenum(j)&"|"
			end if
			next
			votenumlen=len(votenum_1)
			votenum_1=left(votenum_1,votenumlen-1)
			rs("votenum")=votenum_1
			rs("voters")=rs("voters")+1
			rs.update
			conn.execute("update Topic set VoteTotal=voteTotal+"&postnum&" where topicid="&Announceid)
			conn.execute("insert into voteuser (voteid,userid,voteoption) values ("&voteid&","&userid&",'"&postoption&"')")
			end if
	end if
	rs.close
	set rs=nothing

	sql="update topic set LastPostTime=Now() where Topicid="&cstr(announceid)
	conn.execute(sql)
	response.redirect("list.asp?boardid="&boardid&"")
end sub
%>