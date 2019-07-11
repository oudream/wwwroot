<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->

<%if reg="1" then%>
	<html><head>
	<!--#include file=User_Top.asp-->
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title><%=jjgn%> - 服务条款和声明</title>
	<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
	<link rel="stylesheet" type="text/css" href="site.css">
	</head>
	<body topmargin="0">
	<div align="center">
	<form align="center" method="post" name="FrmAddLink" LANGUAGE="javascript"  action="adduser1.asp">

			
			<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="760" id="AutoNumber1">
			<TBODY>
				<TR>
				   <td colspan=2 width="100%" height="25" align="center">
                       ┊ <strong>服 务 条 款 和 声 明</strong> ┊</div></TD>
				</TR>
				<TR>
					<TD>
						<P> 欢迎您加入【<%=jjgn%>】，【<%=jjgn%>】为新闻发布系统，为维护网上公共秩序和社会稳定，请您自觉遵守以下条款：<br><br>
						一、不得利用本站危害国家安全、泄露国家秘密，不得侵犯国家社会集体的和公民的合法权益，不得利用本站制作、复制和传播下列信息： <br><br>
						（一）煽动抗拒、破坏宪法和法律、行政法规实施的； <br>
						（二）煽动颠覆国家政权，推翻社会主义制度的； <br>
						（三）煽动分裂国家、破坏国家统一的； <br>
						（四）煽动民族仇恨、民族歧视，破坏民族团结的； <br>
						（五）捏造或者歪曲事实，散布谣言，扰乱社会秩序的； <br>
						（六）宣扬封建迷信、淫秽、色情、赌博、暴力、凶杀、恐怖、教唆犯罪的； <br>
						（七）公然侮辱他人或者捏造事实诽谤他人的，或者进行其他恶意攻击的； <br>
						（八）损害国家机关信誉的； <br>
						（九）其他违反宪法和法律行政法规的； <br>
						（十）进行商业广告行为的。 <br>
						<br>
						二、互相尊重，对自己的言论和行为负责。 <br>
						</P>
					</TD>
				</TR>
				<TR>
					<TD  align="center">
						<INPUT  type=submit value=我接受条约同意注册 name=Submit>
						<INPUT  onclick="javascript:window.close();" type=button value=我不同意注册 name=Submit1>
					</TD>
				</TR>
			</TBODY>
			</TABLE>
	</form>
	</body>
	</html>
	<!--#include file=User_Bottom.asp-->
<%else%>
	<script language=javascript>
		alert("对不起，用户注册功能已被管理员关闭！")
	</script>
	<body onload="javascript:window.close()"></body>
<%end if%>