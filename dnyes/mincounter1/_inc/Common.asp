<%
	//-- �Լ��������� --//
	//ACCESS��·��
	DbPath				=	"../db_1asdf34/db_CC_Counter6.tfa"

	//ACCESS������(���û�����ÿ�)
	DbPassword			=	""

	//COCOON Counter ��Ȩ�û�
	Accredit			=	"COCOON Counter 6 Free User"

	//����Ա����
	AdminPwd			=	"winer031203"

	//����ԱEmail
	WebMasterEmail		=	"support@dnyes.com"

	//һ��IP���η��������ʱ�䣬��λ�����ӣ�Ĭ��ֵ��20��
	IpTimeExpire		=	20

	//�Ƿ�����ע�ᣬ����ֵ: true/false | 1/0��Ĭ��ֵ: true
	AllowRegister		=	true

	//Ĭ��ÿҳ��ʾ�ļ�¼����Ĭ��: 20
	DefaultPageSize		=	20

	//ʱ�����ã���λ: Сʱ��Ĭ��: 0 - ����������ʱ��
	TimeZoneOffset		=	0
%>

<script language="JScript" runat="server">
	function formatDateTime(d,t){
		function formatNumber(n){return (n+'').length==1?'0'+n:n;}
		var sY=d.getFullYear();sM=d.getMonth()+1,sD=d.getDate();
		var sH=formatNumber(d.getHours()),sN=formatNumber(d.getMinutes());sS=formatNumber(d.getSeconds());
		var sDate = sY+'-'+sM+'-'+sD, sTime = sH+':'+sN+':'+sS;
		switch(t){
			case 0: return sDate + ' ' + sTime;
			case 1: case 2: return sDate;
			case 3:	case 4: return sTime;
		}
	}
</script>

<%
	sConnString			=	"Provider=Microsoft.Jet.OLEDB.4.0;"
						+	"Data Source=" + Server.MapPath(DbPath) + ";"
						+	"Jet OLEDB:Database Password="+DbPassword+";"
	oConn				=	new ActiveXObject("ADODB.Connection");
	oConn.ConnectionString = sConnString;

	sUrl				=	Request.ServerVariables("URL")(1);
%>