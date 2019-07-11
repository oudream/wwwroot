<%
	//-- 自己定义设置 --//
	//ACCESS库路径
	DbPath				=	"../db_1asdf34/db_CC_Counter6.tfa"

	//ACCESS库密码(如果没有请置空)
	DbPassword			=	""

	//COCOON Counter 授权用户
	Accredit			=	"COCOON Counter 6 Free User"

	//管理员密码
	AdminPwd			=	"winer031203"

	//管理员Email
	WebMasterEmail		=	"support@dnyes.com"

	//一个IP两次访问相隔的时间，单位：分钟，默认值：20。
	IpTimeExpire		=	20

	//是否允许注册，允许值: true/false | 1/0，默认值: true
	AllowRegister		=	true

	//默认每页显示的记录数，默认: 20
	DefaultPageSize		=	20

	//时差设置，单位: 小时，默认: 0 - 服务器当地时间
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