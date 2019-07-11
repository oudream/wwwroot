<%
response.Expires=0
%>
<html>
<title>井字游戏</title>
<body>
<Table border=1 align=center>
<%

function test4(m,n)

if m>3 or n>3 or m<0 or n<0 then exit function

if a(m,1)=a(m,2) and a(m,2)=a(m,3) then
     if a(m,1)=1 then
       test4=1
       exit function
     end if
end if

if a(1,n)=a(2,n) and a(2,n)=a(3,n) then
     if a(1,n)=1 then
       test4=1
       exit function
     end if
end if

if a(1,1)=a(2,2) and a(1,1)=a(3,3) then
  if a(2,2)=1 then
    test4=1
    exit function
  end if
end if

if a(1,3)=a(2,2) and a(2,2)=a(3,1) then
  if a(2,2)=1 then
    test4=1
    exit function
  end if
end if

test4=0
end function

function test3(m)
dim i
for i=1 to 3
if a(i,1)=a(i,2) and a(i,3)=0 then
  if a(i,1)=m then
     a(i,3)=2
     test3=1
     exit function
  end if
elseif a(i,2)=a(i,3) and a(i,1)=0 then
  if a(i,2)=m then
     a(i,1)=2
     test3=1
     exit function
  end if  
elseif a(i,1)=a(i,3) and a(i,2)=0 then
  if a(i,1)=m then
     a(i,2)=2
     test3=1
     exit function
  end if
end if
next

for i=1 to 3
if a(1,i)=a(2,i) and a(3,i)=0 then
  if a(1,i)=m then
     a(3,i)=2
     test3=1
     exit function
  end if
elseif a(2,i)=a(3,i) and a(1,i)=0 then
  if a(2,i)=m then
     a(1,i)=2
     test3=1
     exit function
  end if  
elseif a(1,i)=a(3,i) and a(2,i)=0 then
  if a(1,i)=m then
     a(2,i)=2
     test3=1
     exit function
  end if
end if
next

if a(1,1)=a(2,2) and a(3,3)=0 then
  if a(1,1)=m then
     a(3,3)=2
     test3=1
     exit function
  end if
elseif a(1,1)=a(3,3) and a(2,2)=0 then
  if a(1,1)=m then
     a(2,2)=2
     test3=1
     exit function
  end if
elseif a(2,2)=a(3,3) and a(1,1)=0 then
  if a(2,2)=m then
     a(1,1)=2
     test3=1
     exit function
  end if
elseif a(1,3)=a(3,1) and a(2,2)=0 then
  if a(1,3)=m then
     a(2,2)=2
     test3=1
     exit function
  end if
elseif a(1,3)=a(2,2) and a(3,1)=0 then
  if a(2,2)=m then
     a(3,1)=2
     test3=1
     exit function
  end if
elseif a(2,2)=a(3,1) and a(1,3)=0 then
  if a(2,2)=m then
     a(1,3)=2
     test3=1
     exit function
  end if
end if

test3=0
end function


function test2
dim m,n
dim RowArray(10)
dim LineArray(10)
dim Count
dim Rand
Count=0
for m=1 to 3
    for n=1 to 3
		if a(m,n)=0 then
			count=count+1
			LineArray(count)=m
			RowArray(count)=n
        end if
	next
next
if count=0 then
	test2=0 
	exit function
else
	randomize
	Rand=Int(rnd * Count + 1 )
	a(LineArray(Rand),RowArray(Rand))=2
    test2=1
end if
end function

dim a(3,3)
dim over
x=request("X")
y=request("Y")

if x>0 and y>0 and x<4 and y<4 then
  a(x,y)=1  	
elseif x=0 then
 if y=0 then 
  session("a")=a
  session("race")=-1
 else
  session("a")=a
  session("race")=-2
 end if
end if

session("race")=session("race")+1

for j=1 to 3
  for i=1 to 3 
    if session("a")(i,j)=1 then
      a(i,j)=1
    elseif session("a")(i,j)=2 then
      a(i,j)=2    
    end if
  next
next

over=0
if session("race")>=0 then
if test4(x,y)=1 then
  over=1
  response.write "You won"
elseif test3(2)=1 then
  response.write "I won"
  session("race")=session("race")+1
  over=2
elseif test3(1)=1 then
  session("race")=session("race")+1
elseif test2=1 then
  session("race")=session("race")+1
else
  over=3
end if
else
session("race")=0
end if
response.write "&nbsp;"
if session("race")=9 then
  response.write "Game Over"
end if

session("a")=a

for j=1 to 3
  response.write "<tr>"
  for i=1 to 3 
    if session("a")(i,j)=0 then
       if over=0 then 
          response.write "<td><a href=Game.asp?X=" & i & "&Y=" & j & " > <font color=white><b>&nbsp;⊙&nbsp;</b></font></a></td>"
       else
          response.write "<td><font color=white><b>&nbsp;⊙&nbsp;</b></font></td>"
       end if
    elseif session("a")(i,j)=1 then
      response.write "<td><font color=blue><b>&nbsp;★&nbsp;</b></font></td>"
    elseif session("a")(i,j)=2 then
      response.write "<td><font color=red><b>&nbsp;⊙&nbsp;</b></font></td>"
    end if
  next
  response.write "</tr> "
next

%>
</table>
<hr>
<center>重新开始<br>
<a href=Game.asp?X=0&Y=-1>你先</a>
<a href=Game.asp?X=0&Y=0>我先</a>
<%=session("race")%>
</center>