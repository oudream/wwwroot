<%
  function isInteger(para)
       on error resume next
       dim str
       dim l,i
       if isNUll(para) then 
          isInteger=false
          exit function
       end if
       str=cstr(para)
       if trim(str)="" then
          isInteger=false
          exit function
       end if
       l=len(str)
       for i=1 to l
           if mid(str,i,1)>"9" or mid(str,i,1)<"0" then
              isInteger=false 
              exit function
           end if
       next
       isInteger=true
       if err.number<>0 then err.clear
   end function

function IsValidEmail(email)

dim names, name, i, c

'Check for valid syntax in an email address.

IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
   IsValidEmail = false
   exit function
end if
for each name in names
   if Len(name) <= 0 then
     IsValidEmail = false
     exit function
   end if
   for i = 1 to Len(name)
     c = Lcase(Mid(name, i, 1))
     if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
       IsValidEmail = false
       exit function
     end if
   next
   if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
   end if
next
if InStr(names(1), ".") <= 0 then
   IsValidEmail = false
   exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
   IsValidEmail = false
   exit function
end if
if InStr(email, "..") > 0 then
   IsValidEmail = false
end if

end function

function htmlencode2(str)
    dim result
    dim l
    if isNULL(str) then 
       htmlencode2=""
       exit function
    end if
    l=len(str)
    result=""
	dim i
	for i = 1 to l
	    select case mid(str,i,1)
	           case "<"
	                result=result+"&lt;"
	           case ">"
	                result=result+"&gt;"
              case chr(13)
	                result=result+"<br>"
	           case chr(34)
	                result=result+"&quot;"
	           case "&"
	                result=result+"&amp;"
              case chr(32)	           
	                'result=result+"&nbsp;"
	                if i+1<=l and i-1>0 then
	                   if mid(str,i+1,1)=chr(32) or mid(str,i+1,1)=chr(9) or mid(str,i-1,1)=chr(32) or mid(str,i-1,1)=chr(9)  then	                      
	                      result=result+"&nbsp;"
	                   else
	                      result=result+" "
	                   end if
	                else
	                   result=result+"&nbsp;"	                    
	                end if
	           case chr(9)
	                result=result+"    "
	           case else
	                result=result+mid(str,i,1)
         end select
       next 
       htmlencode2=result
   end function
}
%>