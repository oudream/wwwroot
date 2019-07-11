<!--#include file="dbm_adminconn.asp" -->
<SCRIPT language=javascript>
<!--
var linkHeader = "";
var linkTarget = "frmright";
var linkFdrRoot = "";
var imagePath="ffolders/"

var fdrname_root = "根文件夹";

var fdritems = new Array();
// leafgif
var imgsrc_line = imagePath+"l_3_1.gif";
var imgsrc_midblk = imagePath+"l_4_1.gif";
var imgsrc_lastblk = imagePath+"l_3_2.gif";

var imgsrc_lastplus = imagePath+"l_2_2.gif"; //处于折叠状态, 同一级别的最后一项
var imgsrc_lastminus = imagePath+"l_2_2.gif"; //处于展开状态, 同一级别的最后一项

var imgsrc_midplus = imagePath+"l_2_2.gif";  //处于折叠状态, 不是同一级别的最后一项
var imgsrc_midminus = imagePath+"l_2_2.gif"; //处于展开状态, 不是同一级别的最后一项

var img_open = imagePath+"l_2_2.gif";
var img_close = imagePath+"l_2_2.gif";


var spanstr = "";
-->
</SCRIPT>

<SCRIPT src="ffolders/tree.js" type=text/javascript></SCRIPT>

<SCRIPT language=javascript>
<!-- 
function switchstatus(id){
	var item = document.getElementById(id);
	if (item.style.display=='none'){
		item.style.display = '';
	}else{
		item.style.display = 'none';
	}
}
function sclose(id){
	var item = document.getElementById(id);
		item.style.display = 'none';
}
function sopen(id){
	var item = document.getElementById(id);
		item.style.display = '';
}
// -->
</SCRIPT>

<SCRIPT src="ffolders/style.js"></SCRIPT>

<SCRIPT>

<!--
SetCookie( "user", "georong", expdate, "", "163.net" );
SetCookie( "uid", "georong", expdate, "", "163.net" );
//-->
</SCRIPT>

<STYLE type=text/css>BODY {
	FONT-SIZE: 9pt
}
TD {
	FONT-SIZE: 9pt
}
INPUT {
	FONT-SIZE: 9pt
}
SELECT {
	FONT-SIZE: 9pt
}
TEXTAREA {
	FONT-SIZE: 9pt
}
P {
	LINE-HEIGHT: 13pt
}
LI {
	LINE-HEIGHT: 13pt
}
A {
	COLOR: #000000
}
A:link {
	COLOR: #000000; TEXT-DECORATION: underline
}
A:hover {
	COLOR: #555555; TEXT-DECORATION: none
}
</STYLE>

<SCRIPT language=JavaScript>
<!--
var sid='SASPTOMQBynAcRji';
//-->

</SCRIPT>
<%
If session("user_adminlevel")=1 Then
%>
<SCRIPT language=JavaScript>
function sclose1()
{
	sclose('left2');
	sclose('left3');
	sclose('left4');
	sclose('left5');
	sclose('left6');
	sclose('left7');
}
function sopen1()
{
	sopen('left2');
	sopen('left3');
	sopen('left4');
	sopen('left5');
	sopen('left6');
	sopen('left7');
}
</SCRIPT>
<%
elseif session("user_adminlevel")=2 Then
%>
<SCRIPT language=JavaScript>
function sclose1()
{
	sclose('left9');
	sclose('left10');
	sclose('left11');
	sclose('left12');
}
function sopen1()
{
	sopen('left9');
	sopen('left10');
	sopen('left11');
	sopen('left12');
}
</SCRIPT>
<%
elseif session("user_adminlevel")=9 Then
%>
<SCRIPT language=JavaScript>
function sclose1()
{
	sclose('left13');
	sclose('left14');
	sclose('left15');
}
function sopen1()
{
	sopen('left13');
	sopen('left14');
	sopen('left15');
}
</SCRIPT>
<%
end if
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="dbm_css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<TABLE 
width=195 border=0 align="center" cellPadding=0 cellSpacing=0 bgColor=#ffffff style="MARGIN: 0px 0px 0px 9px">
  <TBODY>
  <TR vAlign=top>
      <TD>		
<TABLE width=100% height="30" border=0>
                <TBODY>
                  <TR> 
                    
              <TD height="40"><strong><font color="#FF0000">新筑建筑设计后台管理</font></strong> 
                <br><br> <a href="#" onclick=sclose1();>全关闭</a> | <a href="#" onclick=sopen1();>全打开</a> 
              </TD>
                  </TR>
                </TBODY>
              </TABLE>
      </TD>
    </TR></TBODY></TABLE>
<table width="195" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td>&nbsp;</td>
	    
    <td align="center"><TABLE height=100% width=100% border=0>
        <TBODY>
          <TR> 
            <TD vAlign=top align=left> <br>
<%
If session("user_adminlevel")=1 Then
%>
			  <TABLE width='100%' height=24 align="center"  bordercolor="#CCCCCC" bgcolor="#f5f5f5" style='MARGIN-TOP: 2px'>
          <TBODY>
            <TR> 
              <TD width="68%"><font color="#FF0000"><strong>管理员资料</strong></font></TD>
              <TD width="32%"><IMG style='CURSOR: hand' onclick=switchstatus('left2'); height=24 src='ffolders/l_2_2.gif' width=41 align=absMiddle></TD>
            </TR>
          </TBODY>
        </TABLE>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5" id=left2 style='DISPLAY: none; MARGIN: 6px 0px 0px 0px'>
                <TBODY>
                  <TR> 
                    <TD><a href="dbm_user_edit.asp?user_id=<%=session("user_id")%>" target="mainFrame">自我资料修改</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_user_add.asp" target="mainFrame">管理员 添加</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_user_view.asp" target="mainFrame">管理员 查看修改</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_user_guest.asp" target="mainFrame">管理员 管理的权限</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
			  <TABLE width='100%' height=24 align="center"  bordercolor="#CCCCCC" bgcolor="#f5f5f5" style='MARGIN-TOP: 2px'>
          <TBODY>
            <TR> 
              <TD width="68%"><font color="#FF0000"><strong>客户管理</strong></font></TD>
              <TD width="32%"><IMG style='CURSOR: hand' onclick=switchstatus('left3'); height=24 src='ffolders/l_2_2.gif' width=41 align=absMiddle></TD>
            </TR>
          </TBODY>
        </TABLE>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5" id=left3 style='DISPLAY: none; MARGIN: 6px 0px 0px 0px'>
                <TBODY>
                  <TR> 
                    <TD><a href="dbm_guest_add.asp" target="mainFrame">客户 添加</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_guest_view.asp" target="mainFrame">客户 查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
			  <TABLE width='100%' height=24 align="center"  bordercolor="#CCCCCC" bgcolor="#f5f5f5" style='MARGIN-TOP: 2px'>
          <TBODY>
            <TR> 
              <TD width="68%"><font color="#FF0000"><strong>作品管理</strong></font></TD>
              <TD width="32%"><IMG style='CURSOR: hand' onclick=switchstatus('left4'); height=24 src='ffolders/l_2_2.gif' width=41 align=absMiddle></TD>
            </TR>
          </TBODY>
        </TABLE>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5" id=left4 style='DISPLAY: none; MARGIN: 6px 0px 0px 0px'>
                <TBODY>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_subs_add.asp" target="mainFrame">作品添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_subs_view.asp" target="mainFrame">作品查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
<%
sql="select * from dbm_message" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
zrecordcount=rs.recordcount
rs.close
set rs=nothing
%>					
			  <TABLE width='100%' height=24 align="center"  bordercolor="#CCCCCC" bgcolor="#f5f5f5" style='MARGIN-TOP: 2px'>
          <TBODY>
            <TR> 
              <TD width="68%"><font color="#FF0000"><strong>短信息管理</strong></font></TD>
              <TD width="32%"><IMG style='CURSOR: hand' onclick=switchstatus('left5'); height=24 src='ffolders/l_2_2.gif' width=41 align=absMiddle></TD>
            </TR>
          </TBODY>
        </TABLE>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5" id=left5 style='DISPLAY: none; MARGIN: 6px 0px 0px 0px'>
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000">(您共有<font color="#000000"><%=zrecordcount%></font>条信息) </font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_message_add.asp" target="mainFrame">短信息添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_message_view.asp" target="mainFrame">短信息查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
			  <TABLE width='100%' height=24 align="center"  bordercolor="#CCCCCC" bgcolor="#f5f5f5" style='MARGIN-TOP: 2px'>
          <TBODY>
            <TR> 
                    <TD width="68%"><font color="#FF0000"><strong>留言管理</strong></font></TD>
              <TD width="32%"><IMG style='CURSOR: hand' onclick=switchstatus('left6'); height=24 src='ffolders/l_2_2.gif' width=41 align=absMiddle></TD>
            </TR>
          </TBODY>
        </TABLE>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5" id=left6 style='DISPLAY: none; MARGIN: 6px 0px 0px 0px'>
                <TBODY>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_note_add.asp" target="mainFrame">留言添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_note_view.asp" target="mainFrame">留言查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
			  <TABLE width='100%' height=24 align="center"  bordercolor="#CCCCCC" bgcolor="#f5f5f5" style='MARGIN-TOP: 2px'>
          <TBODY>
            <TR> 
                    <TD width="68%"><font color="#FF0000"><strong>收藏夹管理</strong></font></TD>
              <TD width="32%"><IMG style='CURSOR: hand' onclick=switchstatus('left7'); height=24 src='ffolders/l_2_2.gif' width=41 align=absMiddle></TD>
            </TR>
          </TBODY>
        </TABLE>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5" id=left7 style='DISPLAY: none; MARGIN: 6px 0px 0px 0px'>
                <TBODY>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_hildden_add.asp" target="mainFrame">收藏夹添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_hildden_view.asp" target="mainFrame">收藏夹查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5" id=left8>
                <TBODY>
                  <TR> 
                    <TD width="118" height="35"><a href='folder_manager_check.asp' target='mainFrame'><strong><font color="#FF0000">服务器信息</font></strong></a></TD>
                  </TR>
                  <TR> 
                    <TD width="118" height="35"><a href="folder_manager.asp" target="_blank"><strong><font color="#FF0000">文件系统管理</font></strong></a></TD>
                  </TR>
                </TBODY>
              </TABLE>
<%
End If
%>
<%
If session("user_adminlevel")=2 Then
%>
			  <TABLE width='100%' height=24 align="center"  bordercolor="#CCCCCC" bgcolor="#f5f5f5" style='MARGIN-TOP: 2px'>
          <TBODY>
            <TR> 
              <TD width="68%"><font color="#FF0000"><strong>自我资料管理</strong></font></TD>
              <TD width="32%"><IMG style='CURSOR: hand' onclick=switchstatus('left9'); height=24 src='ffolders/l_2_2.gif' width=41 align=absMiddle></TD>
            </TR>
          </TBODY>
        </TABLE>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5" id=left9 style='DISPLAY: none; MARGIN: 6px 0px 0px 0px'>
                <TBODY>
                  <TR> 
                    <TD><a href="dbm_user_edit.asp?user_id=<%=session("user_id")%>" target="mainFrame">自我资料修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
<%
sql="select * from dbm_message where user_id="&session("user_id")
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
zrecordcount=rs.recordcount
rs.close
set rs=nothing
%>					
			  <TABLE width='100%' height=24 align="center"  bordercolor="#CCCCCC" bgcolor="#f5f5f5" style='MARGIN-TOP: 2px'>
          <TBODY>
            <TR> 
                    <TD width="68%"><font color="#FF0000"><strong>短信息管理</strong></font></TD>
              <TD width="32%"><IMG style='CURSOR: hand' onclick=switchstatus('left10'); height=24 src='ffolders/l_2_2.gif' width=41 align=absMiddle></TD>
            </TR>
          </TBODY>
        </TABLE>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5" id=left10 style='DISPLAY: none; MARGIN: 6px 0px 0px 0px'>
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000">(您共有<font color="#000000"><%=zrecordcount%></font>条信息) </font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_message_add.asp" target="mainFrame">短信息添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_message_view.asp" target="mainFrame">短信息查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
			  <TABLE width='100%' height=24 align="center"  bordercolor="#CCCCCC" bgcolor="#f5f5f5" style='MARGIN-TOP: 2px'>
          <TBODY>
            <TR> 
                    <TD width="68%"><font color="#FF0000"><strong>留言管理</strong></font></TD>
              <TD width="32%"><IMG style='CURSOR: hand' onclick=switchstatus('left11'); height=24 src='ffolders/l_2_2.gif' width=41 align=absMiddle></TD>
            </TR>
          </TBODY>
        </TABLE>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5" id=left11 style='DISPLAY: none; MARGIN: 6px 0px 0px 0px'>
                <TBODY>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_note_add.asp" target="mainFrame">留言添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_note_view.asp" target="mainFrame">留言查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5" id=left12>
                <TBODY>
                  <TR> 
                    <TD width="118" height="35"><a href="folder_manager.asp" target="_blank"><strong><font color="#FF0000">文件系统管理</font></strong></a></TD>
                  </TR>
                </TBODY>
              </TABLE>
<%
End If
%>
<%
If session("user_adminlevel")=9 Then
%>
			  <TABLE width='100%' height=24 align="center"  bordercolor="#CCCCCC" bgcolor="#f5f5f5" style='MARGIN-TOP: 2px'>
          <TBODY>
            <TR> 
              <TD width="68%"><font color="#FF0000"><strong>自我资料管理</strong></font></TD>
              <TD width="32%"><IMG style='CURSOR: hand' onclick=switchstatus('left13'); height=24 src='ffolders/l_2_2.gif' width=41 align=absMiddle></TD>
            </TR>
          </TBODY>
        </TABLE>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5" id=left13 style='DISPLAY: none; MARGIN: 6px 0px 0px 0px'>
                <TBODY>
                  <TR> 
                    <TD><a href="dbm_guest_edit.asp?guest_id=<%=session("user_id")%>" target="mainFrame">自我资料修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
<%
sql="select * from dbm_message where guest_id="&session("user_id")
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
zrecordcount=rs.recordcount
rs.close
set rs=nothing
%>					
			  <TABLE width='100%' height=24 align="center"  bordercolor="#CCCCCC" bgcolor="#f5f5f5" style='MARGIN-TOP: 2px'>
          <TBODY>
            <TR> 
                    <TD width="68%"><font color="#FF0000"><strong>短信息管理</strong></font></TD>
              <TD width="32%"><IMG style='CURSOR: hand' onclick=switchstatus('left14'); height=24 src='ffolders/l_2_2.gif' width=41 align=absMiddle></TD>
            </TR>
          </TBODY>
        </TABLE>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5" id=left14 style='DISPLAY: none; MARGIN: 6px 0px 0px 0px'>
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000"> (您共有<font color="#000000"><%=zrecordcount%></font>条信息) </font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_message_add.asp" target="mainFrame">短信息添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_message_view.asp" target="mainFrame">短信息查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
			  <TABLE width='100%' height=24 align="center"  bordercolor="#CCCCCC" bgcolor="#f5f5f5" style='MARGIN-TOP: 2px'>
          <TBODY>
            <TR> 
                    <TD width="68%"><font color="#FF0000"><strong>留言管理</strong></font></TD>
              <TD width="32%"><IMG style='CURSOR: hand' onclick=switchstatus('left15'); height=24 src='ffolders/l_2_2.gif' width=41 align=absMiddle></TD>
            </TR>
          </TBODY>
        </TABLE>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5" id=left15 style='DISPLAY: none; MARGIN: 6px 0px 0px 0px'>
                <TBODY>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_note_add.asp" target="mainFrame">留言添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_note_view.asp" target="mainFrame">留言查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5" id=left16>
                <TBODY>
                  <TR> 
                    <TD width="118" height="35"><a href="folder_manager.asp" target="_blank"><strong><font color="#FF0000">文件系统管理</font></strong></a></TD>
                  </TR>
                </TBODY>
              </TABLE>
<%
End If
%>
            </TD>
          </TR>
        </TBODY>
      </TABLE> </td>
        <td></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="50" align="center" valign="middle">
          <FORM action="dbm_login.asp" method="post" id="form_logout" name="form_logout" target="_parent">
            <INPUT type=submit value="logout" id="option" name="option" onClick="return confirm('确认退出并关闭此窗口!')">
          </FORM>
	</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>