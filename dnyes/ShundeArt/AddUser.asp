<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->

<%if reg="1" then%>
	<html><head>
	<!--#include file=User_Top.asp-->
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title><%=jjgn%> - �������������</title>
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
                       �� <strong>�� �� �� �� �� �� ��</strong> ��</div></TD>
				</TR>
				<TR>
					<TD>
						<P> ��ӭ�����롾<%=jjgn%>������<%=jjgn%>��Ϊ���ŷ���ϵͳ��Ϊά�����Ϲ������������ȶ��������Ծ������������<br><br>
						һ���������ñ�վΣ�����Ұ�ȫ��й¶�������ܣ������ַ�������Ἧ��ĺ͹���ĺϷ�Ȩ�棬�������ñ�վ���������ƺʹ���������Ϣ�� <br><br>
						��һ��ɿ�����ܡ��ƻ��ܷ��ͷ��ɡ���������ʵʩ�ģ� <br>
						������ɿ���߸�������Ȩ���Ʒ���������ƶȵģ� <br>
						������ɿ�����ѹ��ҡ��ƻ�����ͳһ�ģ� <br>
						���ģ�ɿ�������ޡ��������ӣ��ƻ������Ž�ģ� <br>
						���壩�������������ʵ��ɢ��ҥ�ԣ������������ģ� <br>
						����������⽨���š����ࡢɫ�顢�Ĳ�����������ɱ���ֲ�����������ģ� <br>
						���ߣ���Ȼ�������˻���������ʵ�̰����˵ģ����߽����������⹥���ģ� <br>
						���ˣ��𺦹��һ��������ģ� <br>
						���ţ�����Υ���ܷ��ͷ�����������ģ� <br>
						��ʮ��������ҵ�����Ϊ�ġ� <br>
						<br>
						�����������أ����Լ������ۺ���Ϊ���� <br>
						</P>
					</TD>
				</TR>
				<TR>
					<TD  align="center">
						<INPUT  type=submit value=�ҽ�����Լͬ��ע�� name=Submit>
						<INPUT  onclick="javascript:window.close();" type=button value=�Ҳ�ͬ��ע�� name=Submit1>
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
		alert("�Բ����û�ע�Ṧ���ѱ�����Ա�رգ�")
	</script>
	<body onload="javascript:window.close()"></body>
<%end if%>