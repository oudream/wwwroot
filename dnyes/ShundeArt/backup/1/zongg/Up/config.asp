<%

'*****************************  ���淢�����ڣ�2004-2-18  ����  ��Ȩ˵��  ����  �뱣������Ϣ ****************************************
'***** ��������:����������ϵͳ                               ****                     English Name:ZonGG                      *****
'***** ��ǰ�汾:ZonGG Version 1.1.0.0             ******************************        Powered By:ZonGG Version 1.1.0.0        *****
'***** ��������:����.���� 2003-2004               ***          ****          ***        Do By:Zon.Hizi 2003-2004                *****
'***** ��Ȩ����:����.���� 2003-2004               ******************************        Copyright:Zon.Hizi 2003-2004            *****
'***** ��������:http://www.zon.cn                 *****        ****         ****        URL:http://www.zon.cn                   *****
'***** ����ʵ��:����                              ***   **********************          ����                                    *****
'***** �������ص�ַ:http://zon.cn/down                                    ******        DownLoad:http://zon.cn/down             *****
'***** ����֧������:hizi@zon.cn                                                         E-mail:hizi@zon.cn                      *****
'***** 3721������:����@����                                                             C-mail:����@����                        *****
'************************************************************************************************************************************

%>

<%
''''''''''''''''''����Ϊ��ϵͳ���ơ��汾����������Ϣ��ʹ��������ر����������޸�'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const version="<font color=#ffffff>Powered By:<A href='http://zon.cn/down' target=_blank><font color=#fffff0>ZonGG Version 1.1.0.0 ����������ϵͳ</font></a></font>"


''''''����Ϊ�ײ���Ȩ��Ϣ����Ȩʹ�����޸���ʾ����'''''''''''''''''''''''''''''''''''''''''
Const copyright="<font color=#ffffff>Copyright &copy;2003-2004 <a href='http://www.zon.cn' target='_blank'><font color=#fffff0>���� ZON.CN</font></a> From China.Beijing.Zonvon"

%>
<%


''''''''''''''''''''''''
'''���ݿ����Ӻ���''
''''''''''''''''''''''''''''
Dim Datapath
%>
<!--#include file="conn.asp"-->
<!--#include file="ubbcode.asp"-->
<%
adsdata="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(Datapath)
set adsconn=server.createobject("ADODB.Connection")

'����ҳ��ÿҳ��ʾ�������
const advertlistnumber=24

'�жϵ�ǰ��ʵURL,����/
Function DqUrl()
DqUrl="http://"&Request.ServerVariables("server_name")&Left(Request.ServerVariables("script_name"),InstrRev(Request.ServerVariables("script_name"),"/")-1)
End Function

'�����ʾ������
Function Ggxslx(lx)
Select Case lx
      Case "txt":Ggxslx="���ı�"
      Case "gif":Ggxslx="GIFͼƬ"
      Case "swf":Ggxslx="Flash����"
      Case "dai":Ggxslx="Ƕ�����"
End select
End Function

'����������
Function Ggdklx(lx)
Select Case lx
      Case 0:Ggdklx="�´���"
      Case else:Ggdklx="������"
End select
End Function

'���λ���Ƶ���
Function Ggwm(place)
set adsrs1=server.createobject("adodb.recordset")
adsrs1.open "select * from place where place="&place,adsconn,1,1
if not adsrs1.eof then
Ggwm=adsrs1(1)
else
Ggwm=""
end if
adsrs1.close
Set adsrs1=nothing
End Function

'���λ���ͱ�ʾ���ֵ���
Function Ggwlxsz(place1)
set adsrs1=server.createobject("adodb.recordset")
adsrs1.open "select * from place where place="&place1,adsconn,1,1
if not adsrs1.eof then
Ggwlxsz=adsrs1(2)
else
Ggwlxsz=0
end if
adsrs1.close
Set adsrs1=nothing
End Function

'���λ�������Ƶ���
Function Ggwlx(place)
set adsrs1=server.createobject("adodb.recordset")
adsrs1.open "select * from place where place="&place,adsconn,1,1
if not adsrs1.eof then
Ggwlx=adsrs1(2)
else
Ggwlx=""
end if
adsrs1.close
Set adsrs1=nothing
Select Case Ggwlx
       Case 1:Ggwlx="ҳ��Ƕ��ѭ��"
       Case 2:Ggwlx="������������"
       Case 3:Ggwlx="������������"
       Case 4:Ggwlx="���Ϲ�������"
       Case 5:Ggwlx="�����������"
       Case 6:Ggwlx="�����������"
       Case 7:Ggwlx="ѭ����������"
End select
End Function
'�������λ����ѡ��

Sub Ggwxlxx(place) 'place �����ж�Ĭ��ѡ��
%>
  <select size=1 name=place>
<%
set adsrs1=server.createobject("adodb.recordset")
adsrs1.open "select * from place",adsconn,1,1
do while not adsrs1.eof
%>
<option value="<%=adsrs1(0)%>" <% if adsrs1(0)=place then :response.write "selected":end if%>><%=adsrs1(1)%></option>
 <%adsrs1.movenext
   loop
   adsrs1.close
   Set adsrs1=nothing%>              
  </select> 

<%end sub%>

<%
'���ó��ù��λ���������˵�
Sub Ggwlei(shu) '���ڱ�ʾ���͵���
%>
 <select size=1 name=placelei>
                    <option value=1 <% if shu=1 then%>selected<%end if%>>ҳ��Ƕ��ѭ��</option>
                    <option value=2 <% if shu=2 then%>selected<%end if%>>������������</option>
                    <option value=3 <% if shu=3 then%>selected<%end if%>>������������</option>
                    <option value=4 <% if shu=4 then%>selected<%end if%>>���Ϲ�������</option>
                    <option value=5 <% if shu=5 then%>selected<%end if%>>�����������</option>
                    <option value=6 <% if shu=6 then%>selected<%end if%>>�����������</option>
                    <option value=7 <% if shu=7 then%>selected<%end if%>>ѭ����������</option>
                  </select>
<%end sub%>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''  �жϹ����Ƿ��½  ''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Iflogin()
if session("masterlogin")<>"superadvertadmin" then
%><BR><BR><BR><BR>
<center>

<form method=POST action=login.asp>
       
                          
                    ����Ա��:<input name="username" maxlength="20" size="20">&nbsp; 
                    ��������:<input type="password" name="password" maxlength="16" size="20"> 
                    <input type="submit" name="Submit3" value="������½">
                
      </form>
</center><BR><BR><BR><BR><BR><BR><BR><BR>
   <!--#include file="boot.asp"-->
<%
response.end
end if

end Sub%><%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''�������Ϣ��⺯���������޸ġ�������֣�'''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub addrk()
if request.querystring("job")="add" then

adsconn.open adsdata
dim getname,geturl,getgif,getplace,getwin,getxslei,adsrs,adssql,getclass,getclicks,getshows,gettime,getintro,gethei,getwid

getname = Trim(Request("name"))
geturl = Trim(Request("url"))
getgif = Trim(Request("gif_url"))
getplace =trim(Request("place"))
getwin =trim(Request("window"))
getxslei = trim( Request( "xslei" ))
getclass=trim(Request("class"))
getintro=trim(Request("intro"))
getwid=trim(Request("wid"))
gethei=trim(Request("hei"))

if getxslei="txt" then
getwid=0
gethei=0
end if

if getclass=0 then
getclicks=0
getshows=0
gettime=now()

elseif getclass=1 then
getclicks=trim(request("clicks1"))
getshows=0
gettime=now()

elseif getclass=2 then
getclicks=0
getshows=trim(request("shows2"))
gettime=now()

elseif getclass=3 then
getclicks=0
getshows=0
gettime=trim(request("time3"))

elseif getclass=4 then
getclicks=trim(request("clicks4"))
getshows=trim(request("shows4"))
gettime=now()

elseif getclass=5 then
getclicks=trim(request("clicks5"))
getshows=0
gettime=trim(request("time5"))

elseif getclass=6 then
getclicks=0
getshows=trim(request("shows6"))
gettime=trim(request("time6"))

elseif getclass=7 then
getclicks=trim(request("clicks7"))
getshows=trim(request("shows7"))
gettime=trim(request("time7"))
end if


set adsrs=server.createobject("adodb.recordset")
if  trim(request.querystring("id"))="" then '��������������
adssql="select * from ads"
adsrs.open adssql,adsconn,1,3
adsrs.AddNew
else                                                '������޸Ĺ����
adssql="select * from ads where id="&int(request.querystring("id"))
adsrs.open adssql,adsconn,1,3
end if
adsrs("act") = 1
adsrs("sitename") = getname
adsrs("url") = geturl
adsrs("gif_url") = getgif
adsrs("place") = getplace
adsrs("xslei") = getxslei
adsrs("hei") = gethei
adsrs("wid") = getwid
adsrs("window") = getwin
adsrs("class") = getclass
adsrs("clicks") = getclicks
adsrs("shows") = getshows
adsrs("lasttime") = gettime
adsrs("regtime") = Now()
adsrs("time") = now()
adsrs("intro")=getintro
adsrs.update
adsrs.close
set adsrs=nothing
adsconn.close
set adsconn=nothing
response.redirect "?job=suc&id="&request.querystring("id")

end if
End Sub
%>