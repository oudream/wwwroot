<%

    '****************************************************************************
    '' @����˵��: ����ļ���չ��
    '' @����˵��:  
    '' @����ֵ:   
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
    '' @����˵��: ����Դ�ַ���Str�ĳ���(һ�������ַ�Ϊ2���ֽڳ�)
    '' @����˵��:  - str [string]: Դ�ַ���
    '' @����ֵ:   - [Int] Դ�ַ����ĳ���
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
    '' @����˵��: ��ȡԴ�ַ���Str��ǰLenNum���ַ�(һ�������ַ�Ϊ2���ֽڳ�)
    '' @����˵��:  - str [string]: Դ�ַ���
    '' @����˵��:  - LenNum [int]: ��ȡ�ĳ���
    '' @����ֵ:   - [string]: ת������ַ���
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
    Cutstr=Left(Trim(Str),X)&"��"
   Loop
  End If
 End Function


    '****************************************************************************
    '' @����˵��: ���ַ����е�str�е�HTML������й���
    '' @����˵��:  - str [string]: Դ�ַ���
    '' @����ֵ:   - [string] ת������ַ���
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