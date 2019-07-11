<%
t=(cint(day(date()))*24+cint(hour(time())))*60+cint(minute(time()))
k=0
i=1
y=0
Do While application("zxip"&i)<>""
	if application("zxip"&i)=Request.ServerVariables("REMOTE_ADDR") then
		application("zxsj"&i)=t
		y=1
	end if
	if t-application("zxsj"&i)>9 or t<application("zxsj"&i) then
		k=k+1
	else
		if k>0 then
			application.lock
			application("zxip"&i-k)=application("zxip"&i)
			application("zxsj"&i-k)=application("zxsj"&i)
			application.unlock
		end if
	end if
	if k>0 then
		application.lock
		application("zxip"&i)=""
		application.unlock
	end if
	i=i+1
loop
if y=0 then
	application("zxip"&i)=Request.ServerVariables("REMOTE_ADDR")
	application("zxsj"&i)=t
else
	i=i-1
end if%>