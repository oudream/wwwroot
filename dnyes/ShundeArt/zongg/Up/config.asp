<%

'*****************************  本版发行日期：2004-2-18  ——  版权说明  ——  请保留此信息 ****************************************
'***** 程序名称:忠网广告管理系统                               ****                     English Name:ZonGG                      *****
'***** 当前版本:ZonGG Version 1.1.0.0             ******************************        Powered By:ZonGG Version 1.1.0.0        *****
'***** 开发制作:忠网.黑子 2003-2004               ***          ****          ***        Do By:Zon.Hizi 2003-2004                *****
'***** 版权所有:忠网.黑子 2003-2004               ******************************        Copyright:Zon.Hizi 2003-2004            *****
'***** 忠网域名:http://www.zon.cn                 *****        ****         ****        URL:http://www.zon.cn                   *****
'***** 网络实名:忠网                              ***   **********************          忠网                                    *****
'***** 程序下载地址:http://zon.cn/down                                    ******        DownLoad:http://zon.cn/down             *****
'***** 技术支持信箱:hizi@zon.cn                                                         E-mail:hizi@zon.cn                      *****
'***** 3721中文邮:黑子@忠网                                                             C-mail:黑子@忠网                        *****
'************************************************************************************************************************************

%>

<%
''''''''''''''''''以下为本系统名称、版本、开发者信息，使用者请务必保留，不得修改'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const version="<font color=#ffffff>Powered By:<A href='http://zon.cn/down' target=_blank><font color=#fffff0>ZonGG Version 1.1.0.0 忠网广告管理系统</font></a></font>"


''''''以下为底部版权信息，授权使用者修改显示内容'''''''''''''''''''''''''''''''''''''''''
Const copyright="<font color=#ffffff>Copyright &copy;2003-2004 <a href='http://www.zon.cn' target='_blank'><font color=#fffff0>忠网 ZON.CN</font></a> From China.Beijing.Zonvon"

%>
<%


''''''''''''''''''''''''
'''数据库连接函数''
''''''''''''''''''''''''''''
Dim Datapath
%>
<!--#include file="conn.asp"-->
<!--#include file="ubbcode.asp"-->
<%
adsdata="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(Datapath)
set adsconn=server.createobject("ADODB.Connection")

'管理页面每页显示广告条数
const advertlistnumber=24

'判断当前真实URL,不含/
Function DqUrl()
DqUrl="http://"&Request.ServerVariables("server_name")&Left(Request.ServerVariables("script_name"),InstrRev(Request.ServerVariables("script_name"),"/")-1)
End Function

'广告显示类型名
Function Ggxslx(lx)
Select Case lx
      Case "txt":Ggxslx="纯文本"
      Case "gif":Ggxslx="GIF图片"
      Case "swf":Ggxslx="Flash动画"
      Case "dai":Ggxslx="嵌入代码"
End select
End Function

'广告打开类型名
Function Ggdklx(lx)
Select Case lx
      Case 0:Ggdklx="新窗口"
      Case else:Ggdklx="本窗口"
End select
End Function

'广告位名称调用
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

'广告位类型标示数字调用
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

'广告位类型名称调用
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
       Case 1:Ggwlx="页内嵌入循环"
       Case 2:Ggwlx="上下排列置入"
       Case 3:Ggwlx="左右排列置入"
       Case 4:Ggwlx="向上滚动置入"
       Case 5:Ggwlx="向左滚动置入"
       Case 6:Ggwlx="弹出多个窗口"
       Case 7:Ggwlx="循环弹出窗口"
End select
End Function
'调出广告位下拉选项

Sub Ggwxlxx(place) 'place 用于判断默认选项
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
'调用常用广告位类型下拉菜单
Sub Ggwlei(shu) '用于表示类型的数
%>
 <select size=1 name=placelei>
                    <option value=1 <% if shu=1 then%>selected<%end if%>>页内嵌入循环</option>
                    <option value=2 <% if shu=2 then%>selected<%end if%>>上下排列置入</option>
                    <option value=3 <% if shu=3 then%>selected<%end if%>>左右排列置入</option>
                    <option value=4 <% if shu=4 then%>selected<%end if%>>向上滚动置入</option>
                    <option value=5 <% if shu=5 then%>selected<%end if%>>向左滚动置入</option>
                    <option value=6 <% if shu=6 then%>selected<%end if%>>弹出多个窗口</option>
                    <option value=7 <% if shu=7 then%>selected<%end if%>>循环弹出窗口</option>
                  </select>
<%end sub%>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''  判断管理是否登陆  ''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Iflogin()
if session("masterlogin")<>"superadvertadmin" then
%><BR><BR><BR><BR>
<center>

<form method=POST action=login.asp>
       
                          
                    管理员名:<input name="username" maxlength="20" size="20">&nbsp; 
                    管理密码:<input type="password" name="password" maxlength="16" size="20"> 
                    <input type="submit" name="Submit3" value="立即登陆">
                
      </form>
</center><BR><BR><BR><BR><BR><BR><BR><BR>
   <!--#include file="boot.asp"-->
<%
response.end
end if

end Sub%><%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''广告条信息入库函数（包含修改、添加两种）'''
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
if  trim(request.querystring("id"))="" then '如果是新增广告条
adssql="select * from ads"
adsrs.open adssql,adsconn,1,3
adsrs.AddNew
else                                                '如果是修改广告条
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