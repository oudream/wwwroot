<%
class domainclass
private after	'�������
Private domain   '����
private get_takenhtml
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
   		 TakenHTML = GetURL("http://www.xwhois.com/default.asp?domain=" & Domain,after)
         pageafter = "<pre>"
		 pagebefore = "</pre>"
   end if
   
   if after=".com.cn" or after=".net.cn" or after=".org.cn" or after=".gov.cn" or after=".cn" then
         TakenHTML = GetURL("http://ewhois.cnnic.net.cn/whois?entity=domain&value="&Domain,after)
		 pageafter = "<table border=1 cellspacing=0 cellpadding=2>"
		 pagebefore = "</table>"
     end if
	 
  if after=".�й�" or after=".��˾" or after=".����" then
        TakenHTML = GetURL("http://cwhois.cnnic.net.cn/whois.jsp?entityname=domain&queryinfo=" &Domain,after)
        pageafter = "<table border=1 cellspacing=0 cellpadding=2>"
        pagebefore = "</table>"
    end if
	
else
 
	   TakenHTML = GetURL("http://seal.cnnic.net.cn/queryserver?nameInfo=" & Domain & "&entityName=keyword","ͨ����ַ")
	   pageafter = "<table border=1 cellspacing=0 cellpadding=2>"
	   pagebefore = "</table>"

end if   
  
   tempcontent=TakenHTML
   if TakenHTML="" then
   TakenHTML="��ѯ�д�"
   else
    if instr(TakenHTML,"û�б�ע��")<>0 then
	  TakenHTML="������û�б�ע��"
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
if InStr(TakenHTML,"No match") > 1 or InStr(TakenHTML,"Register system prompt")>1 or InStr(TakenHTML,"not been registered yet")>1 or InStr(TakenHTML,"û���ҵ������Ŀ")>1 or InStr(TakenHTML,"û�в鵽���������ļ�¼")>1 or InStr(TakenHTML,"��Ϣ������")>1 or InStr(TakenHTML,"Not Exist")>1 or InStr(TakenHTML,"NOT FOUND")>1 or  InStr(TakenHTML,"������ѯ����Ϣ������")>1 or  InStr(TakenHTML,"������û�б�ע��")=1 then 
checkdomain=0
else
checkdomain=1
end if 
end Function

Public Property get exsit()
  exsit = checkdomain
End Property

Public Property get getTakenHTML()
  getTakenHTML = get_takenHTML
End Property
 
end class

%>




