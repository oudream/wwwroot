<%
Rem ==========论坛登陆函数=========
Rem 判断用户登陆
response.buffer=true
function chkuserlogin(username,password,usercookies,ctype)
dim rsUser,article,userclass,titlepic
dim userhidden,lastip,UserLastLogin
dim UserGrade,GroupID,ClassSql,FoundGrade
FoundGrade=False
lastip=Request.ServerVariables("REMOTE_ADDR")
userhidden=request.form("userhidden")
if not isnumeric(userhidden) and userhidden="" then userhidden=2
chkuserlogin=false
if ctype=1 then
	sqlstr=" username='"&checkStr(username)&"'"
elseif trim(username)=trim(membername) then
	sqlstr=" userid="&userid&""
else
	sqlstr=" username='"&checkStr(username)&"'"
end if
sql="select userpassword,lockuser,userclass,article,LastLogin,userid,UserGroupID,titlepic from [User] where "&sqlstr&""
set rsUser=conn.execute(sql)
if rsUser.eof and rsUser.bof then
	chkuserlogin=false
else
	if trim(password)<>trim(rsUser(0)) then
		chkuserlogin=false
	elseif rsUser(1)=1 then
		chkuserlogin=false
	else
		userclass=rsUser(2)
		article=rsUser(3)
		UserLastLogin=rsUser(4)
		userid=rsUser(5)
		GroupID=rsUser(6)
		titlepic=rsUser(7)
		if article<0 then article=0
		chkuserlogin=true
	end if
end if
if chkUserLogin then
REM 判断用户等级资料，当用户级别为跟随文章数增长则自动更新等级
REM 自动更新用户数据
set rsUser=conn.execute("select MinArticle from usertitle where usertitle='"&userclass&"'")
if rsUser.eof and rsUser.bof then
	'先判断该组是否有按照文章升级的，也就是MinArticle不是-1的
	set UserGrade=conn.execute("select top 1 usertitle,titlepic,UserGroupID from usertitle where UserGroupID="&GroupID&" and Minarticle<="&article&" and not Minarticle=-1 order by MinArticle desc")
	if not (UserGrade.eof and UserGrade.bof) then
		userclass=UserGrade(0)
		titlepic=UserGrade(1)
		GroupID=UserGrade(2)
		FoundGrade=true
	end if
	if not FoundGrade then
		'该组在等级表中不按照文章升级
		set UserGrade=conn.execute("select top 1 usertitle,titlepic,UserGroupID from usertitle where UserGroupID="&GroupID&" and Minarticle=-1 order by userclass")
		if not (UserGrade.eof and UserGrade.bof) then
			userclass=UserGrade(0)
			titlepic=UserGrade(1)
			GroupID=UserGrade(2)
			FoundGrade=true
		end if
		if not FoundGrade then
		'如果在等级表中未找到相关记录，则使用组名定义等级，采用最高等级用户的图片
		set UserGrade=conn.execute("select top 1 titlepic from usertitle order by MinArticle desc")
		titlepic=UserGrade(0)
		set UserGrade=conn.execute("select title from UserGroups where UserGroupID="&GroupID)
		userclass=UserGrade(0)
		end if
	end if
else
	if rsUser(0)>-1 then
		set UserGrade=conn.execute("select top 1 usertitle,titlepic,UserGroupID from usertitle where Minarticle<="&article&" and not MinArticle=-1 order by MinArticle desc,usertitleid")
		if not (UserGrade.eof and UserGrade.bof) then
			userclass=UserGrade(0)
			titlepic=UserGrade(1)
			GroupID=UserGrade(2)
		end if
	else
		set UserGrade=conn.execute("select usertitle,titlepic,UserGroupID from usertitle where  usertitle='"&userclass&"'")
		if not (UserGrade.eof and UserGrade.bof) then
			userclass=UserGrade(0)
			titlepic=UserGrade(1)
			GroupID=UserGrade(2)
		end if
	end if
end if

select case ctype
case 1
	if datediff("d",UserLastLogin,Now())=0 then
		sql="update [user] set lastlogin=Now(),logins=logins+1,UserLastIP='"&lastip&"',userclass='"&userclass&"',titlepic='"&titlepic&"',UserGroupID="&GroupID&" where userid="&UserID
	else
		sql="update [user] set userWealth=userWealth+"&Forum_user(4)&",userEP=userEP+"&Forum_user(9)&",userCP=userCP+"&Forum_user(14)&",lastlogin=Now(),logins=logins+1,UserLastIP='"&lastip&"',userclass='"&userclass&"',titlepic='"&titlepic&"',UserGroupID="&GroupID&" where userid="&UserID
	end if
case 2
	sql="update [user] set article=article+1,userWealth=userWealth+"&Forum_user(1)&",userEP=userEP+"&Forum_user(6)&",userCP=userCP+"&Forum_user(11)&",lastlogin=Now(),UserLastIP='"&lastip&"',userclass='"&userclass&"',titlepic='"&titlepic&"',UserGroupID="&GroupID&" where userid="&UserID
case 3
	sql="update [user] set article=article+1,userWealth=userWealth+"&Forum_user(2)&",userEP=userEP+"&Forum_user(7)&",userCP=userCP+"&Forum_user(12)&",lastlogin=Now(),UserLastIP='"&lastip&"',userclass='"&userclass&"',titlepic='"&titlepic&"',UserGroupID="&GroupID&" where userid="&UserID
end select
conn.execute(sql)
if founduser and trim(username)<>trim(membername) then
	Response.Cookies("aspsky").path=cookiepath
	Response.Cookies("aspsky")("username")=""
	Response.Cookies("aspsky")("password")=""
	Response.Cookies("aspsky")("userclass")=""
	Response.Cookies("aspsky")("userid")=""
	Response.Cookies("aspsky")("userhidden")=""
	Response.Cookies("aspsky")("usercookies")=""
	conn.execute("delete from online where username='"&membername&"'")
end if
if isnull(usercookies) or usercookies="" then usercookies="0"
select case usercookies
case "0"
	Response.Cookies("aspsky")("usercookies") = usercookies
case 1
   	Response.Cookies("aspsky").Expires=Date+1
	Response.Cookies("aspsky")("usercookies") = usercookies
case 2
	Response.Cookies("aspsky").Expires=Date+31
	Response.Cookies("aspsky")("usercookies") = usercookies
case 3
	Response.Cookies("aspsky").Expires=Date+365
	Response.Cookies("aspsky")("usercookies") = usercookies
end select
Response.Cookies("aspsky")("username") = username
Response.Cookies("aspsky")("userid") = UserID
Response.Cookies("aspsky")("password") = PassWord
Response.Cookies("aspsky")("userclass") = userclass
Response.Cookies("aspsky")("userhidden") = userhidden
rem 清除图片上传数的限制
response.cookies("aspsky")("upNum")=1
Response.Cookies("aspsky").path=cookiepath
end if
set rsUser=nothing
set UserGrade=nothing
end function
%>
