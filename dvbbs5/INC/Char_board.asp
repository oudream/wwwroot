<%
Rem ==========论坛版面函数=========
Rem 判断认证论坛进入用户
function chkboardlogin(boardid,username)
	dim boarduser
	chkboardlogin=false
	if master then
		chkboardlogin=true
	else
		sql="select boarduser from board where boardid="&boardid
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			chkboardlogin=false
		else
			if isnull(rs(0)) or rs(0)="" then
				chkboardlogin=false
			else
			boarduser=split(rs(0), ",")
			for i = 0 to ubound(boarduser)
				if trim(boarduser(i))=trim(username) then
					chkboardlogin=true
					exit for
				end if
			next
			end if
		end if
		rs.close
		set rs=nothing		
	end if
end function
%>
