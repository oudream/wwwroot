<!--#include file="ConnUser.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="char.inc"-->

<script language=javascript>
function CheckFormUserLogin()
{
	if(document.UserLogin.UserName.value=="")
	{
		alert("�������û�����");
		document.UserLogin.UserName.focus();
		return false;
	}
	if(document.UserLogin.Passwd.value == "")
	{
		alert("���������룡");
		document.UserLogin.Passwd.focus();
		return false;
	}
	if(document.UserLogin.verifycode.value == "")
	{
		alert("��������֤�룡");
		document.UserLogin.verifycode.focus();
		return false;
	}
}
</script>
<SCRIPT language="JavaScript" type="text/javascript">
    // Begin morelink
      function morelink(morelink)
      {
        url = 'MoreLink.asp?linktype=1';
        window.open(url,morelink);
      }
    // End morelink-->
 // Begin linkreg
      function linkreg(linkreg)
      {
        url = 'LinkReg.asp';
        window.open(url,linkreg);
      }
    // End linkreg-->
// Begin vote
      function vote(vote)
      {
        url = 'Vote.asp?stype=view';
        window.open(url,vote);
      }
    // End vote-->
// Begin adduser
      function adduser(adduser)
      {
        url = 'AddUser.asp';
        window.open(url,adduser);
      }
    // End adduser-->
// Begin getpwd
      function getpwd(getpwd)
      {
        url = 'GetPwd.asp';
        window.open(url,getpwd);
      }
    // End getpwd-->
</script>

<%if TransShow="on" then%>
<META content=revealTrans(Transition=23,Duration=1.0) http-equiv=Page-Enter>
<META content=revealTrans(Transition=23,Duration=1.0) http-equiv=Page-Exit>
	<SCRIPT language=JavaScript>
	<!--
	//ҳ������л�Ч��
	function transDemo(n) {
	if (document.all && navigator.userAgent.indexOf("Mac")==-1) {
	t=document.all.transmeta;
	t.style.width=document.body.clientWidth;
	t.style.height=document.body.offsetHeight;
	t.style.top=document.body.scrollTop;
	t.style.backgroundColor="#003333";
	t.style.visibility="visible";
	t.filters[0].transition=n;
	setTimeout("transShow()"); // separated to force screen paint
	} else {
	alert("You can view transitions only on Windows IE 4.0 and later.");
	} 
	}
	
	function transShow() {
	t.filters[0].Apply();
	t.style.visibility="hidden";
	t.filters[0].Play();
	}
	//-->
	</SCRIPT>
<%end if%>
<%		'��ȡ��ǰ URL
dim ViewUrl
if Request.ServerVariables("QUERY_STRING")<>"" then
	ViewUrl="http://"& Request.ServerVariables("SERVER_NAME") &""& Request.ServerVariables("url") &"?"& Request.ServerVariables("QUERY_STRING")&""
else
	ViewUrl="http://"& Request.ServerVariables("SERVER_NAME") &""& Request.ServerVariables("url") &""
end if
response.cookies(Forcast_SN)("ViewUrl")=ViewUrl%>

<div id=menuDiv style='Z-INDEX: 2; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>
<%if topbg=1 then %>
	<script src="clearevents.js"></script>
	<noscript><iframe src=*></iframe></noscript>
<%end if%>
<%if R_BG=1 then %>
	<SCRIPT language=JavaScript src="Float.asp"></SCRIPT>
<%end if%>
<div align="center">
<%if showclub=1 then%>
	<TABLE WIDTH=760 BORDER=0 CELLPADDING=0 CELLSPACING=0>
		<TR> 
			<TD width="165"> <IMG SRC="images/fast_01.gif" WIDTH=165 HEIGHT=24 ALT=""></TD>
	<%if Request.cookies(Forcast_SN)("username")="" then%>
		<%
		Function getcode1()
			Dim test
			On Error Resume Next
			Set test=Server.CreateObject("Adodb.Stream")
			Set test=Nothing
			If Err Then
				Dim zNum
				Randomize timer
				zNum = cint(8999*Rnd+1000)
				Session("verifycode") = zNum
				getcode1= Session("verifycode")		
			Else
				getcode1= "<img src=""getcode.asp"">"		
			End If
		End Function
		%>
			<form method="POST" action="ChkLogin.asp" name="UserLogin" onSubmit="return CheckFormUserLogin();">
			<td class="my" align="right" valign="middle" background="images/fast_02.gif">
				<font color=white>�û���</font><input name="UserName" size="8" font face="����" style="font-size:9pt;background-color:#3D3D6E;color:#EAEAF4">
				<font color=white>���룺</font><input type="password" name="Passwd" size="8" font face="����" style="font-size:9pt; background-color:#3D3D6E;color:#EAEAF4">
				<font color=white>���룺</font><input type="text" name="verifycode" size="4" font face="����" style="font-size:9pt; background-color:#3D3D6E;color:#EAEAF4"> 
				<font color=white><b><span><%=getcode1()%></span></b></font>
				<input type="submit" name="Submit" value="��¼" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="��¼ϵͳ"> 
		<%if reg=1 then%>
				&nbsp;<input type="button" name="Submit2" value="ע��" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" onClick="javascript:adduser()" title="ע���»�Ա"> 
		<%end if%>
				&nbsp;<input type="button" name="Submit2" value="����"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" onClick="javascript:getpwd()" title="���������ˣ�">
			</td>
			</form>
	<%else%>
			<td align="right" valign="middle" style="font-size: 9pt;color: #FFFFFF" background="images/fast_02.gif">
			��ӭ�û���<b><%=Request.cookies(Forcast_SN)("UserName")%></b>&nbsp;&nbsp;
			<%if db_Birthday_Select="FeiTium" then		'�Ա��ֶ���ԭ�����ֶ�%>
				<%=Request.cookies(Forcast_SN)("sex")%>
			<%else
				if db_Birthday_Select="Text" then		'�Ա��ֶ���BBS���ı��Ͱ���������
					if Request.cookies(Forcast_SN)("sex")=1 then%>
						��
					<%else
						if Request.cookies(Forcast_SN)("sex")=0 then%>
							Ů
						<%else%>
							����
						<%end if
					end if
				end if
			end if%>
					����Ȩ�ޣ�<%if Request.cookies(Forcast_SN)("KEY")="super" and Request.cookies(Forcast_SN)("purview")="99999" then%><font color="#FFCC00">��������Ա</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="super" and Request.cookies(Forcast_SN)("purview")<>"99999" then%><font color="#FFCC00">ϵͳ����Ա</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="check" then%><font color="#FFCC00">�������Ա</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="selfreg" then%><font color="#FFCC00">ע���û�</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="smallmaster" then%><font color="#FFCC00">С�����Ա</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="bigmaster" then%><font color="#FFCC00">�������Ա</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="typemaster" then%><font color="#FFCC00">��������Ա</font><%end if%>
					���ĵȼ���<%if Request.cookies(Forcast_SN)("KEY")<>"selfreg" then%><font color="#FFCC00">�ڲ���Ա</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="selfreg" and Request.cookies(Forcast_SN)("reglevel")="1" then%><font color="red">��ͨ</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="selfreg" and Request.cookies(Forcast_SN)("reglevel")="2" then%><font color="red">�߼�</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="selfreg" and Request.cookies(Forcast_SN)("reglevel")="3" then%><font color="red">�ؼ�</font><%end if%>
	
	  <a href="Admin_login.asp" class=my>[����]</a>&nbsp;&nbsp;<a href="Exit.asp" class=my>[�˳�]</a> 
      </td>
			<%end if%>
            <TD width="30" background="images/fast_03.gif">&nbsp; </TD>
    </TR>
  </TABLE>
  	  <%end if%>
    <TABLE WIDTH=760 BORDER=0 CELLPADDING=0 CELLSPACING=0>
    <TR> 
      <TD width="165" background="images/fast_04.gif"><a href="<%=logourl%>">
<%if gd1="1" then%>
      <img src="<%=logo%>" width="165" height="86" border="0" align="absmiddle">
<%else%>
        <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="165" height="86" align="absmiddle">
          <param name=movie value='<%=logo%>'>
      <param name=quality value=High>
      <param name="_cx" value="5662">
      <param name="_cy" value="1640">
      <param name="FlashVars" value="-1">
      <param name="Src" value="<%=logo%>">
      <param name="WMode" value="Window">
      <param name="Play" value="-1">
      <param name="Loop" value="-1">
      <param name="SAlign" value>
      <param name="Menu" value="-1">
      <param name="Base" value>
      <param name="Scale" value="ShowAll">
      <param name="DeviceFont" value="0">
      <param name="EmbedMovie" value="0">
      <param name="BGColor" value>
      <param name="SWRemote" value>
	  <param name="wmode" value="transparent">
      <embed src='<%=logo%>' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width="165" height="86" align="absmiddle"></embed></object><%end if%></a></TD>
      <TD width="490" background="images/fast_05.gif"> 
        <div align="center"> 
          <a href="<%=bannerurl%>"> 
      <%if gd2="1" then%>
      <img src="<%=banner%>" width="468" height="60" border="0" align="absmiddle"> 
      <%else%>
      <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="468" height="60" align="absmiddle">
        <param name=movie value='<%=banner%>'>
  <param name=quality value=High>
  <param name="_cx" value="13785">
  <param name="_cy" value="1588">
  <param name="FlashVars" value="-1">
  <param name="Src" value="<%=banner%>">
  <param name="WMode" value="Window">
  <param name="Play" value="-1">
  <param name="Loop" value="-1">
  <param name="SAlign" value>
  <param name="Menu" value="-1">
  <param name="Base" value>
  <param name="Scale" value="ShowAll">
  <param name="DeviceFont" value="0">
  <param name="EmbedMovie" value="0">
  <param name="BGColor" value>
  <param name="SWRemote" value>
  <param name="wmode" value="transparent">
  <embed src='<%=banner%>' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width="468" height="60" align="absmiddle"></embed></object><%end if%></a>
        </div></TD>
      <TD width="105" COLSPAN=3 background="images/fast_06.gif"> 
        <div align="center"> 
          <TABLE cellSpacing=0 cellPadding=2 width=90% align=center border=0>
            <TBODY>
              <TR> 
                <TD class="my" align=middle><font color="#CCFF66">[</font>&nbsp;<a class="my" href="#" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('http://<%=xpurl%>');" >��Ϊ��ҳ</a>&nbsp;<font color="#CCFF66">]</font></TD>
              </TR>
              <TR> 
                <TD class="my" align=middle><font color="#CCFF66">[</font>&nbsp;<a class="my" href="#" onclick=window.external.AddFavorite(location.href,document.title);>�����ղ�</a>&nbsp;<font color="#CCFF66">]</font></TD>
              </TR>
              <TR> 
                <TD class="my" align=middle><font color="#CCFF66">[</font>&nbsp;<a class="my" href="mailto:<%=email%>">��ϵ����</a>&nbsp;<font color="#CCFF66">]</font></TD>
              </TR>
            </TBODY>
          </TABLE>
				</div>
			</TD>
		</TR>
	</TABLE>
  
  	<table width="760" border="0" cellpadding="0" cellspacing="0" bgcolor="#99CDFF">
	</table>
</div>
<table width="760" border="0" align="center" cellpadding="2" cellspacing="0">
  <tr> 
    <td width="30%" background="images/top5.gif">&nbsp;&nbsp;����չ������ϵͳ ��ǰ����<font color=red>
<!--#include file=zx.asp --><font color=red><%=i%></font>
</font>��</td>
    <form method="post" name="myform" action="Result.asp" target="newwindow">
      <td height="22" align="right" background="images/top5.gif"> 
        <%if showsearch=1 then%>
        <%if search="1" then%>
        <select name="action" style="border:1 solid #CCD5F9; font-size:9pt; background-color:#EAEAF4"��size="1">
          <option selected="" value="">ȫ������</option>
          <option value="title">������</option>
          <option value="content">������</option>
          <option value="editor">������</option>
          <option value="about">���ؼ���</option>
        </select> <input type="text" name="keyword" style="font-size:9pt; background-color:#EAEAF4" size="10" value="�ؼ���" onFocus="if (value =='�ؼ���'){value =''}" onBlur="if (value ==''){value='�ؼ���'}" maxlength="50" /> 
        <input type="submit" name="Submit" value="����" /��style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" > 
        <%else%>
        <%
				dim count
				set rs7=server.createobject("adodb.recordset")
				rs7.Source= "select * from "& db_SmallClass_Table &" order by SmallClassID asc"
				rs7.open rs7.Source,conn,1,1
				%>
        <script language="JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
        count = 0		
        do while not rs7.eof 
        %>
subcat[<%=count%>] = new Array("<%=rs7("SmallClassName")%>","<%=rs7("BigClassid")%>","<%=rs7("SmallClassID")%>");
        <%
        count = count + 1
        rs7.movenext
        loop
        rs7.close
		set rs7=nothing
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.SmallClassID.length = 0; 
    document.myform.SmallClassID.options[document.myform.SmallClassID.length] = new Option("ȫ��С��", "");
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.SmallClassID.options[document.myform.SmallClassID.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
</script> <select name="select" style="border:1 solid #CCD5F9; font-size:9pt; background-color:#EAEAF4" size="1">
          <option selected="" value="">ȫ������</option>
          <option value="title">������</option>
          <option value="content">������</option>
          <option value="editor">������</option>
          <option value="about">���ؼ���</option>
        </select> <select name="BigClassid" style="border:1 solid #CCD5F9; font-size:9pt; background-color:#EAEAF4" onChange="changelocation(document.myform.BigClassid.options[document.myform.BigClassid.selectedIndex].value)" size="1">
          <option selected="" value="">ȫ������</option>
          <%         
					set rs8=server.CreateObject("ADODB.RecordSet")
					rs8.Source="select * from " & db_BigClass_Table & " order by BigClassID"
					rs8.Open rs8.Source,conn,1,1
					do while not rs8.eof
						%>
          <option value="<%=rs8("BigClassid")%>"><%=rs8("BigclassName")%></option>
          <%
						rs8.movenext
					loop
					rs8.close
					set rs8=nothing
				        %>
        </select> <select name="SmallClassID" style="border:1 solid #CCD5F9; font-size:9pt; background-color:#EAEAF4">
          <option selected="" value="">ȫ��С��</option>
        </select> <input type="text" name="keyword" style="font-size:9pt; background-color:#EAEAF4" size="10" value="�ؼ���" onFocus="if (value =='�ؼ���'){value =''}" onBlur="if (value ==''){value='�ؼ���'}" maxlength="50" /> 
        <input type="submit" name="Submit" value="����" / style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" > 
        <%end if%>
        <%end if%>
      </td>
    </form>
  </tr>
  </table>
  <table width="760" border="0" align="center" cellpadding="2" cellspacing="0">
    <tr> 
    <td width="30%" bgcolor="#3D3D6E"> <font color="#FFFFFF"> 
      &nbsp;&nbsp;<%Response.Write FormatDateTime(Date(),1) & "  ����"& Mid("��һ����������",WeekDay(Date),1) %>
      </font> </td>
    <td height="22" align="right" bgcolor="#3D3D6E"> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td class="top"> 
            <%
						dim mymenu
						mymenu=basemenu
						Response.Write mymenu
						%>
          </td>
        </tr>
      </table></td>
  </tr>
  <tr align="right" bgcolor="#535396"> 
    <td height="22" colspan="2"> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td class="top"> 
            <!--#include file="Menu1.asp"-->
          </td>
        </tr>
      </table></td>
  </tr>
</table>
