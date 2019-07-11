<%
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_Attach_Table &" where newsid="&newsid
rs.Open rs.source,conn,1,1
do while not rs.eof
	attachname=rs("attachname")
	attachtype=rs("attachtype")
	attachcontent=rs("attachcontent")

	set rs1=server.CreateObject("ADODB.RecordSet")
	rs1.Source="select * from "& db_System_Table &""
	rs1.Open rs1.source,conn,1,1
	uploadtype=rs1("uploadtype")
	rs1.close
	set rs1=nothing
	upload=split(uploadtype,",")
	for i=0 to ubound(upload)
		if attachtype=lcase(trim(upload(i))) then
			filetype=true
			exit for
		else
			filetype=false
		end if
	next

	if filetype=true then
		if attachtype="rar" or attachtype="zip" or attachtype="doc" or attachtype="ppt" or attachtype="xls" or attachtype="mdb" or attachtype="exe" then
			Response.Write "<center><a target='_blank' title='点击这里下载' href='"& FileUploadPath &""& attachname&"."&attachtype&"'><img src='images/"&attachtype&".gif' border=0 ></a><br>"&attachcontent&"</center>"
		end if
		if attachtype="rm" or attachtype="ram" or attachtype="ra" then
			Response.Write "<center>如果要下载，请右击<a class=class href='"& FileUploadPath &""&attachname&"."&attachtype&"'>这里</a>，选择另存为<br>如果不能播放，请点击<a class=class href='http://207.188.7.150/07c46e1b3d2d4c67f406/asia/windows/rp8-cn-setup.exe'>这里下载播放器</a><br><object classid='clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA'   width=352 height=240><param name='CONTROLS' value='ImageWindow'><param name='CONSOLE' value='Clip1'><param name='AUTOSTART' value='-1'><param name=src value='"& FileUploadPath &""&attachname&"."&attachtype&"'></object><br><object classid='clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA'  width=352 height=60><param name='CONTROLS' value='ControlPanel,StatusBar'><param name='CONSOLE' value='Clip1'></object><br>"&attachcontent&"</center>"
		end if
		if attachtype="asf" or attachtype="avi" or attachtype="wmv" or attachtype="mpg" or attachtype="mpeg" or attachtype="mp3" then
			Response.Write "<center>如果要下载，请右击<a class=class href='"& FileUploadPath &""&attachname&"."&attachtype&"'>这里</a>，选择另存为<br>如果不能播放，请点击<a class=class href='http://download.microsoft.com/download/winmediaplayer/wmp71/7.1/W982KMe/CN/mp71.exe'>这里下载播放器</a><br><object classid='clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95'><param name=Filename value='"& FileUploadPath &""&attachname&"."&attachtype&"'><param name='BufferingTime' value='5'><param name='AutoSize' value='-1'><param name='AnimationAtStart' value='-1'><param name='AllowChangeDisplaySize' value='-1'><param name='ShowPositionControls' value='0'><param name='TransparentAtStart' value='1'><param name='ShowStatusBar' value='1'></object><br>"&attachcontent&"</center>"
		end if
		if attachtype="jpg" or attachtype="bmp" or attachtype="gif" or attachtype="png" or attachtype="tif" then
			Response.Write "<center><a target='_blank' title='点击图片看全图' href='"& FileUploadPath &""&attachname&"."&attachtype&"'><img src='"& FileUploadPath &""&attachname&"."&attachtype&"' border=0 onload='javascript:if(this.width>screen.width-333)this.width=screen.width-333'></a><br>"&attachcontent&"</center>"
		end if
		if attachtype="swf" then
			Response.Write "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width=700 height=525><param name=movie value='"& FileUploadPath &""&attachname&"."&attachtype&"'><param name=quality value=high><embed src='"& FileUploadPath &""&attachname&"."&attachtype&"' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=700  height=525></embed></object><br><center>"&attachcontent&"</center>"
		end if
	else
    		if attachtype<>"" then    
    			Response.Write "<center><a target='_blank' title='点击这里下载' href='"& FileUploadPath &""&attachname&"."&attachtype&"'><img src='images/unk.gif' border=0 ></a><br>"&attachcontent&"</center>"
    	end if
end if
rs.movenext
loop
rs.close
set rs=nothing
%>