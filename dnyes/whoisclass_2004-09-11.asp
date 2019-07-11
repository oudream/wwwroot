<%
class whoisclass
private after	'域名后辍
Private domain   '域名
private get_takenhtml
private takenhtml2
private cexploration1

Public Property Let get_domain(ByVal vData)
    domain = vData
End Property

Public Property Let get_after(ByVal vData)
    after = vData
End Property




function bytes2bstr(vin)
strreturn = ""
for i = 1 to lenb(vin)
thischarcode = ascb(midb(vin,i,1))
if thischarcode < &h80 then
strreturn = strreturn & chr(thischarcode)
else
nextcharcode = ascb(midb(vin,i+1,1))
strreturn = strreturn & chr(clng(thischarcode) * &h100 + cint(nextcharcode))
i = i + 1
end if
next
bytes2bstr = strreturn
end function


Private Function getpage()
   Dim pageurl
   Dim temp
   Set temp = Server.CreateObject("microsoft.xmlhttp")
     pageurl ="http://www.cnolnic.com/cgi-bin1/nwhois.cgi?domainname=" & domain
    With temp
     .Open "GET", pageurl, False, "", ""
      .Send
        getpage = .Responsetext
          getpage = bytes2BSTR(.Responsebody)
    End With
     Set temp = Nothing
End Function

Private Function dowith()
Dim lStrURL
    Dim pagebefore
     Dim pageafter
     Dim tempcontent
     Dim pagestart
     Dim pageend
     dim temps
       pageafter = "<PRE>"
       pagebefore = "</PRE>"
       tempcontent = getpage
       pagestart = InStr(1,tempcontent, pageafter,1)
         If pagestart = 0 Then
            dowith = "意外的错误2!"
            d_exsit = 3
            Exit Function
         Else
            pagestart = pagestart + Len(pageafter) + 1
            pageend = InStr(pagestart, tempcontent, pagebefore,1)
              If pageend = 0 Then
                dowith = "意外的错误!"
                d_exsit = 3
                Exit Function
              Else
               pageend = pageend - pagestart '得到内容长度
               tempcontent = Mid(tempcontent, pagestart, pageend)
             End If
         End If
    
      dowith ="<pre>" &tempcontent&"</pre>"
End Function


Function GetURL(url,ext)
    Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
          With Retrieval
          .Open "GET", url, False, "", ""
          .Send
		  if instr(ext,"cn")<>0 then
          GetURL = .responseText
		  else
		  GetURL = .responsebody
		  GetURL= bytes2bstr(GetURL)
		  GetURL=trim(GetURL)
		end if
          End With
    Set Retrieval = Nothing
End Function

private function checkdomain()
if after<>"" then
   if after=".com" or after=".net" or after=".org" or after=".biz" or after=".info" then
   		 TakenHTML = GetURL("http://www-whois.internic.net/cgi/whois?type=domain&whois_nic=" & Domain,after)
         pageafter = "<pre>"
		 pagebefore = "</pre>"
   end if
   
   if after=".com.cn" or after=".net.cn" or after=".org.cn" or after=".gov.cn" or after=".cn" then
         TakenHTML = GetURL("http://ewhois.cnnic.net.cn/whois?entity=domain&value="&Domain,after)
		 pageafter = "<table border=1 cellspacing=0 cellpadding=2>"
		 pagebefore = "</table>"
     end if
	 
  if after=".中国" or after=".公司" or after=".网络" then
        TakenHTML = GetURL("http://cwhois.cnnic.net.cn/whois.jsp?entityname=domain&queryinfo=" &Domain,after)
        pageafter = "<table border=1 cellspacing=0 cellpadding=2>"
        pagebefore = "</table>"
    end if
	
else
 
	   TakenHTML = GetURL("http://seal.cnnic.net.cn/queryserver?nameInfo=" & Domain & "&entityName=keyword","通用网址")
	   pageafter = "<table border=1 cellspacing=0 cellpadding=2>"
	   pagebefore = "</table>"

end if   
  
   tempcontent=TakenHTML
   if TakenHTML="" then
   TakenHTML="查询有错"
   else
    if instr(TakenHTML,"没有被注册")<>0 then
	  TakenHTML="此域名没有被注册"
	  else
       if instr(1,TakenHTML,pageafter,1)<>0 then
          pagestart=InStr(1,tempcontent, pageafter,1)+ Len(pageafter) + 1 
	      pageend=instr(pagestart,tempcontent,pagebefore,1)
	      pageend = pageend - pagestart 
	      TakenHTML = Mid(tempcontent, pagestart, pageend)
		  TakenHTML=pageafter&TakenHTML&pagebefore
      end if
	  
    end if
  end if
  
get_takenhtml=TakenHTML

if InStr(TakenHTML,"No match") > 1 or InStr(TakenHTML,"Register system prompt")>1 or InStr(TakenHTML,"not been registered yet")>1 or InStr(TakenHTML,"没有找到这个条目")>1 or InStr(TakenHTML,"没有查到符合条件的记录")>1 or InStr(TakenHTML,"信息不存在")>1 or InStr(TakenHTML,"Not Exist")>1 or InStr(TakenHTML,"NOT FOUND")>1 or  InStr(TakenHTML,"你所查询的信息不存在")>1 or  InStr(TakenHTML,"此域名没有被注册")=1 then 
checkdomain=0
else
checkdomain=1
end if 
end Function

Public Property get exsit()
  exsit = checkdomain
End Property

Public Property get getTakenHTML()
    if after=".com" or after=".net" or after=".org" or after=".biz" or after=".info" or after=".com.cn" or after=".net.cn" or after=".org.cn" or after=".gov.cn" or after=".cn" then
	 	takenhtml2=dowith
    else
    	takenhtml2=""
    end if
  getTakenHTML =takenhtml2 & get_takenhtml
End Property
 
end class

%>




