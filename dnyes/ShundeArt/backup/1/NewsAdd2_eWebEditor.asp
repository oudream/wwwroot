<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp"-->
<%
aaas1=Request.ServerVariables("SERVER_NAME")
aaas2=Request.ServerVariables("URL")
aaas2=replace(aaas2,"admin/newsadd2.asp","")
Dim strData
Dim intFieldCount
Dim i

intFieldCount = Request.Form("hdnCount")

For i=1 To intFieldCount
	content = content & Request.Form("hdnBigfield" & i)
Next
content=replace(content,"http://"&xpurl&"/"& FileUploadPath,FileUploadPath)
content=replace(content,"http://"&aaas1&aaas2&FileUploadPath,FileUploadPath)

PicUrl=Request.Form("PicUrl")

Set objRegExp = New Regexp
objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "<img.+?>"
strs=trim(content)

'是否偷图
if request.Form("getphoto")="1" then
	Set Matches =objRegExp.Execute(strs)
	For Each Match in Matches
		RetStr = RetStr &getimgs( Match.Value )
	Next 
end if 

function getimgs(str)
	getimgs=""
	Set objRegExp1 = New Regexp
	objRegExp1.IgnoreCase = True
	objRegExp1.Global = True
	objRegExp1.Pattern = "http://.+?"""
	set mm=objRegExp1.Execute(str)
	For Each Match1 in mm
		getimgs=getimgs&"||"&left(Match1.Value,len(Match1.Value)-1)
	next
end function

function getHTTPPage(url)
	on error resume next
	dim http
	set http=server.createobject("MSXML2.XMLHTTP")
	Http.open "GET",url,false
	Http.send()
	if Http.readystate<>4 then 
		exit function
	end if
	getHTTPPage=Http.responseBody
	set http=nothing
	if err.number<>0 then err.Clear 
		end function
		arrimg=split(retstr,"||")
		allimg=""
		newimg=""
		for i=1 to ubound(arrimg)
			if arrimg(i)<>"" and instr(allimg,arrimg(i))<1 then
				fname=""& FileUploadPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&cstr(i&mid(arrimg(i),instrrev(arrimg(i),".")))
				dim geturl,objStream,imgs
				geturl=trim(arrimg(i))
				imgs=gethttppage(geturl)
				Set objStream = Server.CreateObject("ADODB.Stream")
				objStream.Type =1
				objStream.Open
				objstream.write imgs
				objstream.SaveToFile server.mappath(fname),2
				objstream.Close()
				set objstream=nothing
				allimg=allimg&"||"&arrimg(i)
				newimg=newimg&"||"&fname
			end if
		next
		arrnew=split(newimg,"||")
		arrall=split(allimg,"||")
		for i=1 to ubound(arrnew)
			arrnew(i)=replace(arrnew(i),""& FileUploadPath,FileUploadPath)
			strs=replace(strs,arrall(i),arrnew(i))
			arrnew(i)=replace(arrnew(i),FileUploadPath,"")
			if PicUrl=arrall(i) then
				PicUrl=arrnew(i)
			end if
		next
		content=strs
		if left(Picurl,4)="http" then
			fname=""& FileUploadPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&cstr(i&mid(Picurl,instrrev(Picurl,".")))
			dim imgsa
			imgsa=gethttppage(Picurl)
			Set objStream = Server.CreateObject("ADODB.Stream")
			objStream.Type =1
			objStream.Open
			objstream.write imgsa
			objstream.SaveToFile server.mappath(fname),2
			objstream.Close()
			set objstream=nothing
			aassss=Picurl
			Picurl=fname
			Picurl=replace(Picurl,""& FileUploadPath,"")
			aassss1=Picurl
			content=replace(content,aassss,FileUploadPath & Picurl)
		end if
		if Content="" then
			%>
			<script language=javascript>
			history.back()
			alert("请输入文章内容！")
			</script>
			<%
			Response.End
		end if
		if request.cookies(Forcast_SN)("key")="super" then
			if request("viewhtml")<>"" then 
				%>
				<script language=javascript>
				history.back()
				alert("请取消查看HTML源代码后再添加！")
				</script>
				<%
				Response.End
			end if
		end if
		SmallClassid=Request.Form("SmallClassid")
		bigClassid=Request.Form("bigClassid")
		typeid=Request.Form("typeid")
		title=trim(request.form("title"))
		
		if title="" then
			%>
			<script language=javascript>
			history.back()
			alert("请填写文章标题！")
			</script>
			<%
			Response.End
		end if

		if len(title)>100 then
			%>
			<script language=javascript>
			history.back()
			alert("文章标题过长！")
			</script>
			<%
			Response.End
		end if

		if request.Form("goodnews")="1" then
			goodnews=1
		else
			goodnews=0
		end if

		if request.Form("istop")="1" then
			istop=1
		else
			istop=0
		end if

		SpecialID=Request.Form("SpecialID")
		SpecialID2=Request.Form("SpecialID2")
		Author=replace(trim(Request.Form("Author")),"'","''")
		Original=replace(trim(Request.Form("Original")),"'","''")
		about=Request.Form("about")

		if request.Form("PicUrl")="" then
			picnews=0
		else
			picnews=1
		end if

		Picurl=replace(Picurl,FileUploadPath,"")
		editor=request.form("editor")
		if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="bigmaster" or request.cookies(Forcast_SN)("key")="typemaster" then
			checkked=1
		else
			if request.cookies(Forcast_SN)("key")="selfreg" and fabiaocheck="1" then
				checkked=1
			else
				if request.cookies(Forcast_SN)("key")="smallmaster" then
					checkked=1
				else
					checkked=0
				end if
			end if
		end if
		EnCode=Request.Form("EnCode")
		newsrelated=Request.Form("newsrelated")
		newslevel=Request.Form("newslevel")
		title_color=Request.Form("title_color")
		title_type=Request.Form("title_type")
		title_face=Request.Form("title_face")
		title_size=Request.Form("title_size")

		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_News_Table &"" 
		rs.open sql,conn,1,3
		rs.addnew
		rs("title")=title
		rs("Author")=Author
		''库里不添加,改为在 Readnews.asp 中显示时添加,减少库中的代码量
		''content=replace(content,"<IMG","<IMG onmousewheel='return img_zoom(event,this)' onload='javascript:if(this.width>screen.width-333)this.width=screen.width-333'",1,-1,1) 
		''修改上传文件路径标记为[="uploadfile/]
		content=replace(content,"="&chr(34)&FileUploadPath,"="&chr(34)&"uploadfile/",1,-1,1)
		rs("content")=content
		rs("Original")=Original
		rs("goodnews")=goodnews
		rs("picnews")=picnews
		if newslevel="" then
			rs("newslevel")=0
		else
			rs("newslevel")=newslevel
		end if
		rs("istop")=istop
		rs("editor")=editor
		rs("checkked")=checkked
		rs("BigClassid")=BigClassid
		rs("SmallClassid")=SmallClassid
		rs("SpecialID")=SpecialID
		rs("SpecialID2")=SpecialID2
		rs("EnCode")=EnCode
		rs("newsrelated")=newsrelated
		rs("about")=about
		rs("picname")=PicUrl
		if title_type="" or title_type="0" then
			rs("titletype")="l"
		else
			rs("titletype")=title_type
		end if
		rs("titlecolor")=title_color
		
		if title_face="" then
			rs("titleface")="无"
		else
			rs("titleface")=title_face
		end if

		rs("titlesize")=title_size
		rs("typeid")=typeid
		rs("UpdateTime")=now()
		rs.update
		newsid=rs("newsid")
		dim username
		username=request.cookies(Forcast_SN)("username")
		set rs2=server.createobject("adodb.recordset")
		sql2="select * from "& db_User_Table &" where  "& db_User_Name &"='"&username&"'"
		rs2.open sql2,ConnUser,1,3
		rs2("number")=rs2("number")+1
		rs2.update
		rs2.close
		set rs2=nothing
		
		rs.Close
		set rs1=nothing

		set rs3=server.createobject("adodb.recordset")
    		sql3="select * from "& db_UploadPic_Table &" where username='"&request.cookies(Forcast_SN)("username")&"'" 
    		rs3.open sql3,conn,1,3
    		do while not rs3.eof
    			set rs4=server.createobject("adodb.recordset")
			sql4="select * from "& db_Attach_Table &"" 
           		rs4.open sql4,conn,1,3
           		filename=rs3("picname")
           		rs4.addnew
           		rs4("filename")=filename
           		rs4("newsid")=newsid
           		rs4.update
           		rs4.close
           		set rs4=nothing
           		RS3.MoveNext
           	loop
           	conn.execute("delete from "& db_UploadPic_Table &" where username='"&request.cookies(Forcast_SN)("username")&"'")
           	rs3.close
           	set rs3=nothing
 
		for i=1 to ubound(arrnew)
			set rs4=server.createobject("adodb.recordset")
           		sql4="select * from "& db_Attach_Table &"" 
           		rs4.open sql4,conn,1,3
           		filename=arrnew(i)
           		rs4.addnew
           		rs4("filename")=filename
           		rs4("newsid")=newsid
           		rs4.update
           		rs4.close
           		set rs4=nothing
		next
 
		set rs5=server.createobject("adodb.recordset")
           	sql5="select * from "& db_Attach_Table &"" 
           	rs5.open sql5,conn,1,3
           	filename=aassss1
           	rs5.addnew
           	rs5("filename")=filename
           	rs5("newsid")=newsid
           	rs5.update
           	rs5.close
           	set rs4=nothing
  
		Response.cookies(Forcast_SN)("content")=""
		Response.cookies(Forcast_SN)("newsrelated")=""
		response.write "<p align=center><font color=red>恭喜您，文章“"&title&"”已经成功添加!<br>"&freetime&"秒钟后返回上页!</font>"
		
		if request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster" or request.cookies(Forcast_SN)("key")="smallmaster" then
			response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=typemanage.asp"">"
		else
			if smallclassid<>"" then
				response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=listnews.asp?smallclassid="&smallclassid&""">"
			else
				response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=smallno.asp?bigclassid="&bigclassid&""">"
			end if
		end if
		%>