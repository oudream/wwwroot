<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Config.asp"-->
<%
''主程序开始部分

dim formsize,formdata,Msg
formsize = Request.TotalBytes
formdata = Request.BinaryRead(formsize)
UploadSize=True

If formsize = 0 or Formsize > OKsize Then
UploadSize=False
Response.Write"你要上传的文件大小超出程序限制,请<a href=index.asp>返回</a>修改重试"
Response.End
End If


dim sinfo_Stream
Set Sinfo_Stream = Server.CreateObject("adodb.stream")
Sinfo_Stream.Type = 1		''2进制流
Sinfo_Stream.Mode = 3		''读写模式
Sinfo_Stream.Open
Sinfo_Stream.Write formdata		''保存二进制内容到流对象
''分离数据变量
dim VbEnter
dim spStr,lenOfspStr,bpos
dim loopcnt,exitflag,ppoint,npoint
''保存数据变量		
dim FldData,fldHeadStr,infldpos
dim databpos,datalen
dim FldInfo(15,1)
''fldInfo(0)表单头内容
''fldInfo(1)表单数据

VbEnter = chrb(13)&chrb(10)''读取第一个VbEnter位置
bpos = Instrb(formdata,VbEnter)
SpStr = midb(formdata,1,bpos+1) ''包含了一个0d0a
LenOfspStr = lenb(Spstr) 
ppoint = LenOfspStr+1 ''位置指针,指向每一个表单域内容的开始位置
formdata = midb(formdata,ppoint)
loopcnt = 0   ''表单元素
do 
	bpos = instrb(formdata,spStr) ''分割位置
	npoint = (ppoint+bpos+lenofspstr-1)  ''指向下一表单开始位置
	if bpos < 1 then
		fldData = midb(formdata,1,instrb(formdata,leftb(spStr,lenOfspstr-2))-1)
		bpos = lenb(fldData)+1
		exitflag = true
	else
		FldData = leftb(formdata,bpos-1)		
		formdata = midb(formdata,bpos+LenOfspstr)
	end if
	infldpos = instrb(fldData,vbEnter&vbEnter)
	fldHeadStr = bytes2bstr(midb(fldData,1,infldpos-1))
	fldInfo(loopcnt,0) = fldHeadStr	''表单头
	''Response.Write fldHeadStr&"<br>"
	databpos = (ppoint+infldpos-1+4)
	Sinfo_Stream.Position = databpos-1
	datalen = (bpos-infldpos-6)
	if datalen = 0 then
		fldInfo(loopcnt,1) = ""
	else
		fldInfo(loopcnt,1) = Sinfo_Stream.Read(datalen)
	end if
	ppoint = npoint
	loopcnt = loopcnt + 1
loop until exitflag = true
Sinfo_Stream.close
Set Sinfo_Stream = Nothing


''以上程序数据处理过程
''写入数据库并处理文件上传开始
Sub SaveData()

ftitle = MyRequest("filetitle")
Msg = ""
		if ftitle = "" then 
			Msg = Msg & "文件名称：空<br>"
		else
			Msg = Msg & "文件名称："&ftitle&"<br>"
		end if
		ftype = myrequest("fileType")		
		Msg = Msg & "文件类型："&ftype&"<br>"
		filedata = myrequest("filedata")
		filesize = lenb(filedata)
		if  filesize = 0 then 
			Msg = Msg & "上传文件：没有<br>"
		else 
			filename = GetFileName("filedata")
			''限制加入的类型 *.asp
			file_ctype = GetContentType("filedata")
			Msg = Msg & "上传文件："&filename&"&nbsp;&nbsp;&nbsp;"
			Msg = Msg & "数据流："&file_ctype&"&nbsp;&nbsp;&nbsp;"
			Msg = Msg & "文件长度："&filesize&"<br>"
		end if
		filedesc = myrequest("fileDesc")
		Msg = Msg & "文件说明："&filedesc&"<br><br>"
		
		FileTypeName = GetFileTypeName(FileName)
		If  IsvalidFile(FileTypeName)=False Then
		Msg = "文件类型非法，不允许上传"&FileTypeName&"文件！"
		Exit Sub
		End If		

		if ftitle<>"" and fileSize > 0 and UploadSize=True then
			''保存数据到数据库
			dim basepath,sql
			basepath = "./subsimg/"
			sql = "insert into info (filetitle,filedesc,filetype,filecontenttype,filepath,filesize) values ('"
			sql = sql & ftitle &"','"&filedesc&"','"&ftype&"','"&file_ctype&"','"&basepath&filename&"',"&filesize&")"
			myconn.Execute(sql)
			Call SavetoFile(filedata,basepath,filename)
			Msg = Msg & "文件已经上传<br>"
		else
			Msg = Msg & "上传失败！ "&ErrorMsg&"<br>"
		end if
		myconn.close()
		set myconn = nothing

End Sub	
''文件上传已经写入数据完毕，提示信息出口为变量msg
SaveData
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ok</title>
<meta http-equiv="refresh" content="10;URL=Index.asp">
<link href="upstyle.css" rel="stylesheet" type="text/css">
</head>

<body background="../images/mainbg.gif" leftmargin="0" topmargin="0">
<table width="550" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr> 
    <td height="150"> <table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr class="text"> 
          <td><strong>提示信息：</strong></td>
        </tr>
        <tr class="text"> 
          <td> <%=msg%> </td>
        </tr>
        <tr class="text"> 
          <td><div align="center">本页将在10秒后返回 如果您的浏览器没有反应，请<a href=Index.asp>点击此处返回</a></div></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
