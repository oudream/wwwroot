<%

    '****************************************************************************
    '' @功能说明: 获得文件扩展名
    '' @参数说明:  
    '' @返回值:   
    '****************************************************************************

function getFileExtName(fileName)
	dim pos
	pos=instrrev(filename,".")
	if pos>0 then
		getFileExtName=mid(fileName,pos+1)
	else
		getFileExtName=""
	end if
end function

function gotTopic(str,strlen)
	dim l,t,c
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i)&""
			exit for
		else
			gotTopic=str&""
		end if
	next
end function

    '****************************************************************************
    '' @功能说明: 计算源字符串Str的长度(一个中文字符为2个字节长)
    '' @参数说明:  - str [string]: 源字符串
    '' @返回值:   - [Int] 源字符串的长度
    '****************************************************************************
 Public Function strLen(Str)
  If Trim(Str)="" Or IsNull(str) Then 
   strlen=0
  else
   Dim P_len,x
   P_len=0
   StrLen=0
   P_len=Len(Trim(Str))
   For x=1 To P_len
    If Asc(Mid(Str,x,1))<0 Then
     StrLen=Int(StrLen) + 2
    Else
     StrLen=Int(StrLen) + 1
    End If
   Next
  end if
 End Function

    '****************************************************************************
    '' @功能说明: 截取源字符串Str的前LenNum个字符(一个中文字符为2个字节长)
    '' @参数说明:  - str [string]: 源字符串
    '' @参数说明:  - LenNum [int]: 截取的长度
    '' @返回值:   - [string]: 转换后的字符串
    '****************************************************************************
 Public Function CutStr(Str,LenNum)
  Dim P_num
  Dim I,X
  If StrLen(Str)<=LenNum Then
   Cutstr=Str
  Else
   P_num=0
   X=0
   Do While Not P_num > LenNum-2
    X=X+1
    If Asc(Mid(Str,X,1))<0 Then
     P_num=Int(P_num) + 2
    Else
     P_num=Int(P_num) + 1
    End If
    Cutstr=Left(Trim(Str),X)&"…"
   Loop
  End If
 End Function


    '****************************************************************************
    '' @功能说明: 将字符串中的str中的HTML代码进行过滤
    '' @参数说明:  - str [string]: 源字符串
    '' @返回值:   - [string] 转换后的字符串
    '****************************************************************************

function nohtml(str)
    dim re
    Set re=new RegExp
    re.IgnoreCase =true
    re.Global=True
    re.Pattern="(\<.[^\<]*\>)"
    str=re.replace(str," ")
    re.Pattern="(\<\/[^\<]*\>)"
    str=re.replace(str," ")
    nohtml=str
    set re=nothing
end function

 %>