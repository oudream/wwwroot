<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="function.asp"-->

<%
IF request.cookies(Forcast_SN)("KEY")="" THEN

else
	usernamecookie=CheckStr(request.cookies(Forcast_SN)("UserName"))
	passwdcookie=CheckStr(trim(Request.cookies(Forcast_SN)("passwd")))
	KEYcookie=CheckStr(trim(request.cookies(Forcast_SN)("KEY")))
	if usernamecookie="" or passwdcookie="" then
		response.cookies(Forcast_SN)("UserName")=""
		response.cookies(Forcast_SN)("KEY")=""
		response.cookies(Forcast_SN)("purview")=""
		response.cookies(Forcast_SN)("fullname")=""
		response.cookies(Forcast_SN)("reglevel")=""
	else
		'�ж��û��ĺϷ���
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_User_Table &" where "& db_User_Name &"='"&usernamecookie&"'"
		rs.open sql,ConnUser,1,1
		if rs.eof and rs.bof then
			response.cookies(Forcast_SN)("UserName")=""
			response.cookies(Forcast_SN)("KEY")=""
			response.cookies(Forcast_SN)("purview")=""
			response.cookies(Forcast_SN)("fullname")=""
			response.cookies(Forcast_SN)("reglevel")=""
		end if
		IF passwdcookie<>rs(db_User_Password) THEN
			response.cookies(Forcast_SN)("UserName")=""
			response.cookies(Forcast_SN)("KEY")=""
			response.cookies(Forcast_SN)("purview")=""
			response.cookies(Forcast_SN)("fullname")=""
			response.cookies(Forcast_SN)("reglevel")=""
		END IF
		'�����ж��û�����ʵ�������û������Ƕ�Ӧ���ж�
		if KEYcookie<>rs("OSKEY") then
			response.cookies(Forcast_SN)("UserName")=""
			response.cookies(Forcast_SN)("KEY")=""
			response.cookies(Forcast_SN)("purview")=""
			response.cookies(Forcast_SN)("fullname")=""
			response.cookies(Forcast_SN)("reglevel")=""
		end if
		rs.close
		set rs=nothing
	END IF
END IF



'���ļ���Ҫ���е���������
dim typename
NewsID=Request.QueryString("NewsID")
Page=Request.QueryString("page")

if page="" then
	page=1
	elseif not IsNumeric(page) then
		Show_Err("�Ƿ�������<br><br><a href='javascript:history.back()'>����</a>")
		response.end
	end if
	page=int(page)
	if newsid="" then
		Show_Err("δָ��������<br><br><a href='javascript:history.back()'>����</a>")
		response.end
	elseif not IsNumeric(newsid) then
		Show_Err("�Ƿ�������<br><br><a href='javascript:history.back()'>����</a>")
		response.end
	else
		'�жϸ�ƪ�����Ƿ����
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_News_Table &" where NewsId="&NewsId
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			rs.close
			set rs=nothing
			Show_Err("ָ�������²����ڣ�<br><br><a href='javascript:history.back()'>����</a>")
			response.end
		else
			checked=rs("checkked")
			rs.close
			set rs=nothing
		end if
		
		if checked=1 or Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" then	'��������˲�������ع���Ա
			conn.execute("update "& db_News_Table &" Set Click=click+1 where NewsID=" & NewsID )
		end if
		set rs=server.CreateObject("ADODB.RecordSet")
		if uselevel=1 then
			if Request.cookies(Forcast_SN)("key")="" then
				rs.Source="select * from "& db_News_Table &" where checkked=1 and newslevel=0 and newsid="&newsid
			end if
			if Request.cookies(Forcast_SN)("key")="selfreg" then
				if Request.cookies(Forcast_SN)("reglevel")=3 then
					rs.Source="select * from "& db_News_Table &" where checkked=1 and newslevel<=3 and newsid="&newsid
				end if
				if Request.cookies(Forcast_SN)("reglevel")=2 then
					rs.Source="select * from "& db_News_Table &" where checkked=1 and newslevel<=2 and newsid="&newsid
				end if
				if Request.cookies(Forcast_SN)("reglevel")=1 then
					rs.Source="select * from "& db_News_Table &" where checkked=1 and newslevel<=1 and newsid="&newsid
				end if
			end if
			if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
				rs.Source="select * from "& db_News_Table &" where newsid="&newsid
			else
				rs.Source="select * from "& db_News_Table &" where newsid="&newsid
			end if
			rs.Open rs.Source,conn,1,1
			bigclassid=rs("bigclassid")
			smallclassid=rs("smallclassid")
			title=htmlencode4(trim(rs("title")))
			title1=htmlencode4(trim(rs("title")))
			about=htmlencode4(trim(rs("about")))
			Author=htmlencode4(trim(rs("Author")))
			editor=htmlencode4(trim(rs("editor")))
			Original=htmlencode4(trim(rs("Original")))
			image=rs("image")
			UpdateTime=trim(rs("UpdateTime"))
			News_Content=rs("Content")
			SpecialID=rs("SpecialID")
			SpecialID2=rs("SpecialID2")
			click=rs("click")
			EnCode=trim(rs("EnCode"))
			typeid=rs("typeid")
			titletype=rs("titletype")
			titlecolor=rs("titlecolor")
			titleface=rs("titleface")
			editor=rs("editor")
			wzdj=rs("newslevel")
			backtype=rs("backtype")
			rs.Close
			set rs=nothing

			set rs=server.CreateObject("ADODB.RecordSet")
			rs.Source="select * from "& db_Type_Table &" where typeID=" & typeID
			rs.Open rs.Source,conn,1,1
			typename=rs("typename")
			rs.Close
			set rs=nothing
			set rs=server.CreateObject("ADODB.RecordSet")
			rs.Source="select * from "& db_BigClass_Table &" Where BigClassid=" & BigClassid
			rs.Open rs.Source,conn,1,1
			bigclassname=rs("bigclassname")
			rs.close
			set rs=nothing
			if smallclassid<>"" then
				set rs=server.CreateObject("ADODB.RecordSet")
				rs.Source="select * from "& db_SmallClass_Table &" Where smallClassid=" & smallClassid
				rs.Open rs.Source,conn,1,1
				smallclassname=rs("smallclassname")
				rs.close
				set rs=nothing
			end if%>
<html><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=title%><%if smallclassid<>"" then%>_<%=SmallClassName%><%end if%>_<%=BigClassName%>_<%=typename%>_<%=jjgn%></title>

<%if backtype=0 then %>
<LINK href=news.css rel=stylesheet>
<% else %>
<LINK href=News_1.css rel=stylesheet>
<% end if %>

<style type="text/css">
.newstitle  {COLOR: #000000; FONT-FAMILY:"Verdana, Arial, ����"; FONT-SIZE: 14px;line-height:1.5}
</style>

<script language="JavaScript" type="text/JavaScript">

function validateMode()
{
  if (!bTextMode) return true;
  alert("����ȡ���鿴HTMLԴ���룬���롰�༭��״̬��Ȼ����ʹ��ϵͳ�༭����!");
  message.focus();
  return false;
}
function validateModea()
{
  if (!bTextMode) return true;
  alert("����ȡ���鿴HTMLԴ����!");
  message.focus();
  return false;
}

function sign()
{if (!validateMode()) return;
message.focus();

var range =message.document.selection.createRange();
str1="<font color='red'><b>|ǩ��|</b>�ļ��Ѿ��Ķ�</font>"
   range.pasteHTML(str1);
}

function img_zoom(e, o)		//ͼƬ����������
{
	var zoom = parseInt(o.style.zoom, 10) || 100;
	zoom += event.wheelDelta / 12;
	if (zoom > 0) o.style.zoom = zoom + '%';
		return false;
}
</script>

<script language=javascript>
function user_login()
{
	document.UserLogin.UserName.focus();
	return false;
}
</script>

<script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
	var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
	var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
	if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
</script>

<script language="JavaScript" type="text/JavaScript">

var size=14;			//�����С

function fontZoom(fontsize){	//��������
	size=fontsize;
	document.getElementById('fontZoom').style.fontSize=size+'px';
}

function fontMax(){	//����Ŵ�
	size=size+2;
	document.getElementById('fontZoom').style.fontSize=size+'px';
}

function fontMin(){	//������С
	size=size-2;
	if (size < 2 ){
		size = 2
	}
	document.getElementById('fontZoom').style.fontSize=size+'px';
}

</SCRIPT>

<script language="javascript">
<!--
function changedata() {
	if ( document.AddReview.none.checked ) {
		document.AddReview.Author.value = "�ο�";
	} 
}
function changedata1() {
	if ( document.AddReview.none1.checked ) {
		document.AddReview.email.value = "guest@feitium.net";
	} 
}

//-->
</script>

<script language=javascript>
function CheckFormAddReview()	//������۷�������д��Ŀ�Ƿ�Ϊ��
{
	if(document.AddReview.Author.value=="")
	{
		alert("�������û�����");
		document.AddReview.Author.focus();
		return false;
	}
	if(document.AddReview.email.value == "")
	{
		alert("����������EMAIL��");
		document.AddReview.email.focus();
		return false;
	}
	if(document.AddReview.content.value == "")
	{
		alert("�������������ݣ�");
		return false;
	}
}
</script>
</head>
<%if ScrollClick = "double" then%>
	<SCRIPT language=JavaScript>
	//˫����������
	var currentpos,timer;
	
	function initialize()
	{
	timer=setInterval("scrollwindow()",50);
	}
	
	function sc(){
	clearInterval(timer);
	}
	
	function scrollwindow()
	{
	currentpos=document.body.scrollTop;
	window.scroll(0,++currentpos);
	if (currentpos != document.body.scrollTop)
		sc();
	}
	document.onmousedown=sc
	document.ondblclick=initialize
	</script>
<%else%>
	<SCRIPT>
	//�����϶���Ļ��ʽ
	var old_y=0;  //��¼��갴��ʱ�ģ���λ��
	var yn=0;  //���ڼ�¼���״̬
	function moveit()
	{
	if(yn==1 &&  event.button==1)  //������������¾��ƶ�ҳ��
	document.body.scrollTop=(old_y-event.clientY); //�ƶ�ҳ��
	}
	function downit()
	{old_y=event.clientY+document.body.scrollTop; //��ס��갴��ʱ�ģ���λ��
	yn=1; //���ڼ�ס����Ѱ���
	}
	
	function upit()
	{yn=0;}  //��ס����ѷſ�
	
	document.onmouseup=upit; //���ſ�ʱִ��upit()����
	document.onmousedown=downit; //��갴��ʱִ��downit()����
	document.onmousemove =moveit; //����ƶ�ʱִ��moveit()����
	</SCRIPT>
<%end if%>
<body topmargin="0" marginheight="0">
<SCRIPT language=javascript>
<!--
  function do_color(vobject,vvar)
  { document.getElementById(vobject).style.color=vvar; }
  function do_zooms(vobject,vvar)
  { document.getElementById(vobject).style.fontsize=vvar+'px'; }
-->
</SCRIPT>
<!--#include file="top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr valign="top"> 
		<td>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr> 
					<td valign="top">
						<table width="102%" border="0" cellspacing="0" cellpadding="0">
							<tr bgcolor="#FFFFFF"> 
								<td height="3" bgcolor="#FFFFFF"><img src="images/kb.gif" width="9" height="3"></td>
							</tr>
							<tr bgcolor="#FFFFFF"> 
								<td width="100%" height="25" background="images/menu-guestbook.gif" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��Ŀ����<a class=daohang href="./" >&nbsp;</a><strong><a class=daohang href="./" >��վ��ҳ</a>&gt;&gt;<a class=daohang href="Type.asp?TypeId=<%=typeid%>"><%=typename%></a>&gt;&gt;<a class=daohang href="BigClass.asp?TypeId=<%=typeid%>&BigClassid=<%=BigClassid%>" ><%=BigClassName%></a> 
<%if smallclassid<>"" then%>
									&gt;&gt;<a class=daohang href='SmallClass.asp?typeid=<%=typeid%>&BigClassID=<%=BigClassID%>&SmallClassID=<%=SmallClassID%>'><%=SmallClassName%></a> 
<%end if%>
<%title1=htmlencode4(title1)%>
									</strong>
								</td>
							</tr>
							<tr bgcolor="#FFFFFF"> 
	                					<td height="25" background="images/menu-guest-l.gif" bgcolor="#FFFFFF" align="right">
	                 						&nbsp;&nbsp;���� <font color=red><%=click%></font> λ���߶�������&nbsp;&nbsp;
		������ɫ��<select name=do_color_frm size=1 onchange="if(this.options[this.selectedIndex].value!='')
		{do_color('fontZoom',this.options[this.selectedIndex].value);}">
		<option>ѡ����ɫ</option>
		<option style="background-color:Black;color:Black" value=Black>��  ɫ</option>
		<option style="background-color:Red;color:Red" value=Red>��  ɫ</option>
		<option style="background-color:Yellow;color:Yellow" value=Yellow>��  ɫ</option>
		<option style="background-color:Green;color:Green" value=Green>��  ɫ</option>
		<option style="background-color:Orange;color:Orange" value=Orange>��  ɫ</option>
		<option style="background-color:Purple;color:Purple" value=Purple>��  ɫ</option>
		<option style="background-color:Blue;color:Blue" value=Blue>��  ɫ</option>
		<option style="background-color:Brown;color:Brown" value=Brown>��  ɫ</option>
		<option style="background-color:Teal;color:Teal" value=Teal>ī  ��</option>
		<option style="background-color:Navy;color:Navy" value=Navy>��  ��</option>
		<option style="background-color:Maroon;color:Maroon" value=Maroon>��  ʯ</option>
		<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF">�� ��</option>
		<option style="background-color:#7FFFD4;color: #7FFFD4" value="#7FFFD4">�� ��</option>
		<option style="background-color:#FFE4C4;color: #FFE4C4" value="#FFE4C4">�� ��</option>
		<option style="background-color:#7FFF00;color: #7FFF00" value="#7FFF00">�� ��</option>
		<option style="background-color:#D2691E;color: #D2691E" value="#D2691E">�� ��</option>
		<option style="background-color:#FF7F50;color: #FF7F50" value="#FF7F50">ש ��</option>
		<option style="background-color:#6495ED;color: #6495ED" value="#6495ED">�� ��</option>
		<option style="background-color:#DC143C;color: #DC143C" value="#DC143C">�� ��</option>
		<option style="background-color:#FF1493;color: #FF1493" value="#FF1493">õ���</option>
		<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF">�� ��</option>
		<option style="background-color:#FFD700;color: #FFD700" value="#FFD700">�� ��</option>
		<option style="background-color:#DAA520;color: #DAA520" value="#DAA520">�� ��</option>
		<option style="background-color:#808080;color: #808080" value="#808080">�� ��</option>
		<option style="background-color:#778899;color: #778899" value="#778899">�� ��</option>
		<option style="background-color:#B0C4DE;color: #B0C4DE" value="#B0C4DE">�� ��</option>
 		</select>&nbsp;&nbsp;
									�����壺<A class=black href="javascript:fontMax()">�Ŵ�</A>&nbsp;<A class=black href="javascript:fontZoom(14)">����</A>&nbsp;<A class=black href="javascript:fontMin()">��С</A>��&nbsp;&nbsp;&nbsp;&nbsp;
								</td>
							</tr>
	              					<tr bgcolor="#FFFFFF">
								<td height="25" align="right" background="images/menu-l-guest.gif" bgcolor="#FFFFFF"><font color="#000000"><%if ScrollClick = "double" then%>��˫���������Զ�������<%else%>�������������϶�������<%end if%>��ͼƬ�Ϲ��������ֱ佹ͼƬ��&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr> 
					<td valign="top" background="images/menu-guest-l.gif">&nbsp; </td>
				</tr>
			</table>
		</td>
	</tr>
	<tr valign="top"> 
		<td> 
			<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
				<tr> 
          				<td>
          					<table border="0" cellpadding="0" cellspacing="0" width="98%" align="center">
							<tr> 
								<td width="100%" align=center>
									<br>
									<font color="<%=titlecolor%>" size="+2" face="����_GB2312"><strong><%=title1%></strong></font>
									<HR>
									<br>
									&nbsp;&nbsp;�������ڣ�<%=year(updateTime)%>��<%=month(updateTime)%>��<%=day(updateTime)%>��&nbsp;&nbsp; 
	<%if Original<>"" then%>
									������<%=Original%> 
	<%end if%>
									&nbsp;&nbsp; 
	<%if Author<>"" then%>
									���ߣ�<%=Author%> 
	<%end if%>
									&nbsp;&nbsp;&nbsp;&nbsp;���༭¼�룺<a href="User.asp?user=<%=editor%>"><%=editor%></a>��
								</td>
							</tr>
	<%
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_News_Table &" where NewsID=" & NewsID
rs.Open rs.source,conn,1,1
typeid=rs("typeid")
Title=trim(rs("Title"))
image=rs("image")

dim mode
    
set rs11=server.CreateObject("ADODB.RecordSet")
rs11.Source="select * from "& db_Type_Table &" where typeid="&typeid&" order by typeid"
rs11.Open rs11.Source,conn,1,1
mode=rs11("mode")
rs11.close
set rs11=nothing

''���ͼƬ����������Ч��
if mouse_wheel_zoom="on" then
	News_Content=replace(News_Content,"<IMG","<IMG onmousewheel='return img_zoom(event,this)' onload='javascript:if(this.width>screen.width-333)this.width=screen.width-333'",1,-1,1) 
end if
''ͼƬ�ϴ�·����ԭΪ config.asp ���趨�� [FileUploadPath] ֵ
News_Content=replace(News_Content,"="&chr(34)&"uploadfile/","="&chr(34)&FileUploadPath,1,-1,1)

arr_Content=split(News_Content,"[---��ҳ---]")
MaxPages=ubound(arr_Content)
%>
							<tr> 
								<td width="100%"  align="center">
									<table border="0" cellspacing="0" cellpadding="0" align="center" style="TABLE-LAYOUT: fixed">
										<tr> 
											<td width="100%" align=center></td>
										</tr>
										<tr> 
											<TD class=newstitle id=fontzoom vAlign=top><br>
                        
<%if M_BG=1 and rs("picnews")=0 and Not Instr(rs("Content"),"TD")>0 then%>
												<table border=0 align="left" cellpadding=3>
													<tr> 
														<td>
															<script language=javascript src="zongg/ad.asp?i=15"></script>
														</td>
													</tr>
												</table>
<%end if%>

<%if (checked<>1) and ((Request.cookies(Forcast_SN)("key")<>"super") and (Request.cookies(Forcast_SN)("key")<>"typemaster") and (Request.cookies(Forcast_SN)("key")<>"bigmaster") and (Request.cookies(Forcast_SN)("key")<>"smallmaster")) then	'����δ���,�����Ƿ���ع���Ա
	response.write "<P><CENTER><strong><font color='0000ff' size='+2' face='����'>���»�δ�������<br>��ȴ�����֪ͨ����Ա��˲���������</font></strong></CENTER>"
	response.write "<P><CENTER><IMG SRC='" &ReadNews_CopyRight_Logo& "' BORDER='0' ALT=''></CENTER>"
else	'���������
	if checked<>1 then
		response.write "<P><CENTER><strong><font color='ff00ff' size='+2' face='����'>���ѣ������»�δͨ�����</font></strong></CENTER>"
	end if
	if cINT(wzdj)<1 then
		Response.Write arr_Content(Page-1)%>
		<CENTER><IMG SRC="<%=ReadNews_CopyRight_Logo%>" BORDER="0" ALT=""></CENTER>
	<% else %>
		<%if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then %>
			<%Response.Write arr_Content(Page-1)%>
			<CENTER><IMG SRC="<%=ReadNews_CopyRight_Logo%>" BORDER="0" ALT=""></CENTER>
		<% else %>
			<%if  Request.cookies(Forcast_SN)("key")="" then%>
				<br><font color="#cc0000"><b>���ݼ�飺</b></font><br><br>
				<%=CutStr(nohtml(rs("Content")),150)%>... 
				<br><br>
				<br><font color="#cc0000"><b>�������ѣ�</b></font><br><br>
				�㻹ûע�᣿����û�е�¼������Ȩ�޲�������ƪ����Ҫ���Ǳ�վ����Ȩ��Ҫ���ע���û������Ķ���<br><br>
				<%
				response.write "���¼���"
				response.write cINT(wzdj)
				response.write "��"
				%>
				<br>
				<%
				response.write "����Ȩ�ޣ�"
				response.write "δע��"
				%>
				<br><br>
				����㻹ûע�ᣬ��Ͻ����<a href="#" OnClick="adduser()"  class=bottom><font color="#cc0000">ע��</font></a>�ɣ�<br>
				<br>
				������Ѿ�ע�ᵫ��û��¼����Ͻ����<a href="#" OnClick="user_login()"  class=bottom><font color="#cc0000">��¼</font></a>�ɣ�<br>
				<CENTER><IMG SRC="<%=ReadNews_CopyRight_Logo%>" BORDER="0" ALT=""></CENTER>
			<% else %>
				<%if  Request.cookies(Forcast_SN)("key")="selfreg" then
					if cINT(Request.cookies(Forcast_SN)("reglevel"))>=cINT(wzdj) then%>
						<%Response.Write arr_Content(Page-1)%>
						<CENTER><IMG SRC="<%=ReadNews_CopyRight_Logo%>" BORDER="0" ALT=""></CENTER>
					<% else %>
						<br><font color="#cc0000"><b>���ݼ�飺</b></font>
						<br><br>
						<%=CutStr(nohtml(rs("Content")),150)%>... 
						<br><br>
						<br><font color="#cc0000"><b>�������ѣ�</b></font><br><br>
						�㻹ûע�᣿����û�е�¼������Ȩ�޲�������ƪ����Ҫ���Ǳ�վ����Ȩ��Ҫ���ע���û������Ķ���<br><br>
						<%
						response.write "���¼���"
						response.write cINT(wzdj)
						response.write "��"
						%><br>
						<%
						response.write "����Ȩ�ޣ�"
						response.write (Request.cookies(Forcast_SN)("reglevel"))
						response.write "��"
						%>
						<br><br>
						����㻹ûע�ᣬ��Ͻ����<a href="#" OnClick="adduser()" class=bottom><font color="#cc0000">ע��</font></a>�ɣ�<br>
						<br>
						������Ѿ�ע�ᵫ��û��¼����Ͻ����<a href="#" OnClick="user_login()"  class=bottom><font color="#cc0000">��¼</font></a>�ɣ�<br>
						<CENTER><IMG SRC="<%=ReadNews_CopyRight_Logo%>" BORDER="0" ALT=""></CENTER>
						<br>
					<%end if%>
				<%end if%>
			<%end if%>
		<%end if%>
	<%end if%>
<%end if%>
												<br>
												<div align="right">
<%
url="ReadNews.asp?NewsId="&newsid
if MaxPages >0 then
	Response.write "<a class=black href='"& Url &"&page=1' title='��1ҳ'>��ҳ</a> "
	if Page-1 > 0 then
		Prev_Page = Page - 1
		Response.write "<a class=black href='"& Url &"&page="& Prev_Page &"' title='��"& Prev_Page &"ҳ'>��һҳ</a> "
	end if

	for PageCounter=0 to MaxPages
		PageLink = PageCounter+1
		if PageLink <> Page Then
			Response.write "<a class=black href='"& Url &"&page="& PageLink &"'>["& PageLink &"]</a> "
		else
			Response.Write "<font color='#FF0000'><B>["& PageLink &"]</B></font> "
		end if
		If PageLink = MaxPages+1 Then Exit for
	Next
	if page <= Maxpages then
		bdd_Page = Page + 1
		Response.write "<a class=black href='" & Url & "&page=" & bdd_Page & "' title='��" & bdd_Page & "ҳ'>��һҳ</A>"
	end if
	Response.write " <A class=black href='" & Url & "&page=" & Maxpages+1 & "' title='��"& Maxpages+1 &"ҳ'>βҳ</A>"
end if
%>
												</div>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr> 
								<td width="100%"  height="25">
									<div align="center"> 
									<!--#include file=attach.asp -->
									</div>
								</td>
							</tr>
							<tr> 
								<td width=100% ><hr size=1></td>
							</tr>
							<tr> 
								<td height=8></td>
							</tr>
							<tr> 
								<td width=100% height=18><B>���ר�⣺</b> 
<%set rs4=server.CreateObject("ADODB.RecordSet")
if SpecialID<>0 then
	rs4.Source="select * from "& db_Special_Table &" where SpecialID=" & SpecialID
	rs4.Open,conn,1,3
	if not rs4.eof then%>
		<a class=class href='Special_News.asp?SpecialID=<%=SpecialID%>'><%=trim(rs4("specialname"))%></a> 
	<%end if
	rs4.Close
	set rs4=nothing

	set rs4=server.CreateObject("ADODB.RecordSet")
	if specialid2<>0 then
		rs4.Source="select * from "& db_Special_Table &" where SpecialID=" & SpecialID2
		rs4.Open,conn,1,3
		if not rs4.eof then %>
			<a class=class href='Special_News.asp?SpecialID=<%=SpecialID2%>'><%=trim(rs4("specialname"))%></a> 
		<%end if
		rs4.Close
		set rs4=nothing
	end if%>
								</td>
							</tr>
							<tr> 
								<td width=100% height=18><B>ר����Ϣ��</b></td>
							</tr>
							<tr> 
								<td height=8></td>
							</tr>
	<%
	set rs=server.CreateObject("ADODB.RecordSet")
	rs.Source="select top 5 * from "& db_News_Table &" where NewsID<>"& NewsID & " and (SpecialID=" & SpecialID & " or SpecialID2=" & SpecialID & ") and checkked=1 order by NewsID DESC"
	rs.Open,conn,1,1
	
	if rs.EOF then
		Response.Write "<tr><td width=100% >&nbsp;û��ר����Ϣ</td></tr>"
	else
		while not rs.EOF
			title=htmlencode4((rs("title")))
			Response.Write "<tr><td width=100% height=18>&nbsp;<img src=images/goxp.gif>&nbsp;<a class=middle target=_top href='ReadNews.asp?NewsID=" & rs("NewsID") & "'>" & Title & "</a><font color=#666666>(" & trim(rs("UpdateTime")) &")[<font color=#ff0000>" & rs("click") & "</font>]</font></td></tr>"
			rs.MoveNext
		wend
		Response.Write "<tr><td width=98% align=right height=18><a class=lift href='Special_News.asp?SpecialID=" & SpecialID &"'><img src='images/more.gif' border='0'></a></td></tr>"
			
	end if

	else
	Response.Write "<font color=red ><b>ר����Ϣ��</b></font>"

	
	rs.Close
	set rs=nothing
end if
'------------------------------------------------------------------------------------------------------------------------------------
Response.Write "<tr><td width=100% ><hr size=1></td></tr><tr><td height=8></td></tr>"
if about<>""  then
Response.Write "<tr><td width=100% height=18><B>�����Ϣ:</b></td></tr><tr><td height=8></td></tr>"
	sql="select top 5 * from "& db_News_Table &" where (about like '%" & about & "%' or title like '%" & about & "%') and checkked=1 order by newsid desc"
	set rs=conn.execute(sql)

	do while not rs.eof
		title=htmlencode4(trim(rs("title")))
		%>
							<tr> 
								<td height=18> <img src=images/goxp.gif> <a class=middle target=_top href='ReadNews.asp?NewsID=<%=rs("NewsID")%>'><%=Title%></a><font color=#666666>(<%=month(trim(rs("updateTime")))%>��<%=day(trim(rs("updateTime")))%>��)[<font color=#666666><%=rs("click")%></font>]</font></td>
							</tr>
		<% rs.movenext
	loop
	Response.Write "<tr><td width=98% align=right height=18><a class=lift href='Result.asp?keyword=" & about &"&action=title'><img src='images/more.gif' border='0'></a></td></tr>"
	rs.close   
	set rs=nothing  
else
Response.Write "<tr><td width=100% height=18><B><font color=red >�����Ϣ��</font></b></td></tr><tr><td height=8></td></tr>"

end if

set rs1=server.CreateObject("ADODB.RecordSet")
rs1.Source="select top "&reviewnum&" * from "& db_Review_Table &" where NewsID=" & NewsID & " and passed=1 order by reviewid desc"
rs1.Open rs1.Source,conn,1,1
if rs1.EOF then  NoReview=1
	Response.Write "<tr><td width=100% ><hr size=1></td></tr><tr><td height=8></td></tr>"
	Response.Write "<tr><td width=100% ><B>�������:</b></td></tr><tr><td height=8></td></tr>"
	%>
							<tr> 
								<td width="100%"> 
	<%
	if NoReview then 						
		Response.Write "<font color=red ><b>���������</b></font>"
	end if
	%>
								</td>
							</tr>
	<%
	if not NoReview then
		while not rs1.EOF
			author=server.HTMLEncode(trim(rs1("author")))
			email=server.HTMLEncode(trim(rs1("email")))
			reviewip=rs1("reviewip")
			content=trim(rs1("content"))
            content=replace(content,"table","�������")
			content=replace(content,"TABLE","�������")
			''content=replace(content,"<","&lt;")
			''content=replace(content,">","&gt;")
			''content=replace(content,chr(13),"<BR>")
			ContentLen=len(Content)
			%>
							<tr> 
								<td>
									<table width="100%" border="0" cellspacing="0" cellpadding="5" style="table-layout:fixed; word-break:break-all">
										<tr bgcolor="#EFEFEF"> 
											<td width="322">�����ˣ�<%=author%></td>
											<td width="270"> 
												<p align="right"> 
			<%if Request.cookies(Forcast_SN)("key")="super" or showip="1" then%>
												IP��<%=reviewip%> 
			<%end if%>
											</td>
										</tr>
										<tr> 
											<td colspan="2">
												<table width="100%" border="0" cellspacing="0" cellpadding="0" style="table-layout:fixed; word-break:break-all">
													<tr> 
														<td>�������ʼ���<a href='mailto:<%=email%>'><%=email%></a></td>
														<td align="right">����ʱ�䣺<%=rs1("updatetime")%></td>
													</tr>
												</table>
											</td>
										</tr>
										<tr bgcolor="#FFFFFF">
											<td colspan="2"><%
			If ContentLen=<50  then	
				DisplayContent=Content
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;" & displaycontent
										%></td>
										</tr><%
			else
									%><tr bgcolor="#FFFFFF"> 
											<td colspan="2"><%
				DisplayContent=replace(nohtml(trim(Content)),"&nbsp;","",1,-1,1) '��ȡ����������ֶ����ݲ��滻��ʽ��
				DisplayContent=replace(DisplayContent,vbcrlf,"",1,-1,1)
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"& CutStr(displaycontent,50) &"<a href='ReadView.asp?ReviewID=" & rs1("ReviewID") &"&NewsID=" & NewsID &"' target=_blank class=class>��ϸ����</a>"
			end if
										%></td>
										</tr>
									</table>
								</td>
							</tr><%
			rs1.MoveNext
		wend
	end if

	rs1.Close
	set rs1=nothing	
						%><tr> 
								<td width="98%" height="28"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>

	<%if reviewable="1" then%>  
		<form method="POST" action="AddReview.asp" name=AddReview  onSubmit="return CheckFormAddReview();">
	                <%if Request.cookies(Forcast_SN)("username")<>"" then
		
				%>
				 <tr> 
					<td width="100%" background="images/menu-news.gif" align="center">
						<table border="0" cellspacing="0" width="54%" cellpadding="4">
							<input type=hidden name=NewsID value=<%=NewsID%>>
							<input type=hidden name=url value="http://<%= Request.ServerVariables("SERVER_NAME") %><%=Request.ServerVariables("url")%>?<%=Request.ServerVariables("QUERY_STRING")%>">
							<tr> 
								<td width="26%" align="left"><a name="send"><font color="#FF0000">*</font>��&nbsp;��&nbsp;����</a></td>
								<td width="70%"><input type="text" name="Author" size="20" value="<%=Request.cookies(Forcast_SN)("UserName")%>" readonly></td>
							</tr>
							<tr> 
								<td width="26%" align="left"><font color="#FF0000">*</font>�����ʼ���</td>
								<td width="70%"><input type="text" name="email" size="20" value="<%=Request.cookies(Forcast_SN)("UserEmail")%>" ></td>
							</tr>
							<tr> 
								<td width="26%" align="left" valign="top"><font color="#FF0000">*</font>�������ݣ�</td>
								<td width="70%" align="left">��100�����ڣ�<% if M_MAIN=1 then %><font color=red >ǩ���ļ�</font>��
								<img class=None  src="images/watermark.gif" align="absmiddle" border="0" style="cursor:hand;" alt="ǩ���ļ�" lANGUAGE="javascript" onclick="sign()"></td><% end if%>
							</tr>
							<tr>
								<td width="70%" colspan="2" align="center" valign="top">
									<script language="javascript">
									var bTextMode=false;
									document.write ('<iframe src="guestbox.asp?action=new" id="message" width="350" height="100" frameborder=yes scrolling=no align=left></iframe>')
									frames.message.document.designMode = "On";
									</script>
								</td>
							</tr>
							<tr> 
								<td width="70%" colspan="2" align="center" height="50"> 
									<input type="hidden" name="editor" value="<%=editor%>"> 
									<input name="passed" type="hidden" value="<%if reviewcheck="1" then%>1<%else%>0<%end if%>"> 
									<input type="submit" value="��������" name="Submit" OnClick="document.AddReview.content.value = frames.message.document.body.innerHTML;">
									<input name="reset" type=reset value="������д"><input type="hidden" name="content" value="">
									<input type="hidden" name="title" value="���ۣ�<%=title%>">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			<%else%>
				<tr> 
					<td width="100%" background="images/menu-news.gif" align="center">
						<table border="0" cellspacing="0" width="54%" cellpadding="4">
							<input type=hidden name=NewsID value=<%=NewsID%>>
							<input type=hidden name=url value="http://<%= Request.ServerVariables("SERVER_NAME") %><%=Request.ServerVariables("url")%>?<%=Request.ServerVariables("QUERY_STRING")%>">
							<tr> 
								<td width="26%" align="right"><a name="send"><font color="#FF0000">*</font>��&nbsp;��&nbsp;����</a></td>
								<td width="70%"> <input type="text" name="Author" size="20">&nbsp;�οͣ�<INPUT onclick=changedata() type=checkbox value=true  name=none></td>
							</tr>
							<tr> 
								<td width="26%" align="right"><font color="#FF0000">*</font>�����ʼ���</td>
								<td width="70%"> <input name="email" type="text" value="">&nbsp;�οͣ�<INPUT onclick=changedata1() type=checkbox value=true  name=none1></td>
							</tr>
							<tr> 
								<td width="26%" align="right" valign="top"><font color="#FF0000">*</font>�������ݣ�</td>
								<td width="70%" align="left">��100�����ڣ�</td>
							</tr>
							<tr>
								<td width="70%" colspan="2" align="center" valign="top">
									<script language="javascript">
									document.write ('<iframe src="guestbox.asp?action=new" id="message" width="350" height="100" frameborder=yes scrolling=no align=left></iframe>')
									frames.message.document.designMode = "On";
									</script>
								</td>
							</tr>
							<tr> 
								<td width="70%" colspan="2" align="center" height="50"> 
									<input type="hidden" name="editor" value="<%=editor%>"> 
									<input name="passed" type="hidden" value="<%if reviewcheck="1" then%>1<%else%>0<%end if%>"> 
									<input type="submit" value="��������" name="Submit" OnClick="document.AddReview.content.value = frames.message.document.body.innerHTML;">
									<input name="reset" type=reset value="������д"><input type="hidden" name="content" value="">
									<input type="hidden" name="title" value="���ۣ�<%=title%>">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			<%end if%>
		</form>
	<%end if%>
    
	<tr valign="top">
		<td height="25" background="images/menu-l-guest.gif"></td>
	</tr>
</table>

<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="bottom">
	 <tr> 
		<td background="images/menu-guest-l.gif">
			<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr> 
					<td width="20"></td>
					<td  width="255" height="20"> 
						<a href=Review.asp?NewsID=<%=NewsID%> target="_blank" class=bottom><img src="images/icon1.gif"  border="0">�����鿴������ڸ���Ϣ������</a>
					</td>
					<td  width="214" height="20"> 
						<a href=send.asp?NewsID=<%=NewsID%> target=_blank  class=bottom><img src="images/mail.gif"  border="0">������Ϣ��������</a>
					</td>
					<td  width="168"><a href="javascript:window.print()" class=bottom><img src="images/printer.gif"  border="0">��ӡ��ҳ</a></td>
					<td  width="90"><input type="button" name="close2" value="�رմ���"  onClick="window.close();return false;"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="19" background="images/menu-guest-t.gif"></td>
	</tr>
</table>
<!--#include file=bottom.asp -->
</body>
</html>
	<%end if%>
<%end if
conn.close
set conn=nothing
%> 