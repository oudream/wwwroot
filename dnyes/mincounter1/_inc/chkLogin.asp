<!--#include file="Common.asp"-->
<script language="javascript" runat="server">
	function chkLogin(b){
		if(Request.QueryString("UID").Count>0){
			if(doLogin()) return true;
			else doAlert("","Login.asp?innerUrl="+sUrl);
		}
		iSiteId = Session("cc_count_6_site_id");
		if(isNaN(iSiteId)||iSiteId<1){
			// not logined
			if(b){
				// do login
				if(doLogin()) return true;
				else doAlert("","Login.asp?innerUrl="+sUrl);
			}else{
				iSiteId = 0;
				return false;
			}
		}else{
			return true;
		}
	}

	function logout(){
		Session.contents.remove("cc_count_6_site_id");
		Session.contents.remove("cc_count_6_admin");
		doAlert("","../supervise/login.asp");
	}

	function doLogin(){
		var sUID = Str4Sql(Request("UID")+"");
		var sPWD = Str4Sql(Request("PWD")+"");
		var sInnerUrl = Str4Sql(Request("innerUrl")+"");
		if(sUID.length<1||sUID=="undefined") return false;
		sSqlString = "select SiteId,pwd,SiteOpen from t_Site where UID='"+sUID+"'";
		var oRs = oConn.execute(sSqlString);
		if(oRs.EOF) doAlert("用户不存在","Login.asp?innerUrl="+sUrl);
		else if(sPWD!=oRs.fields.item(1).value&&oRs.fields.item(2).value<1) doAlert("密码错误","Login.asp");
		iSiteId = oRs.fields.item(0).value;
		Session("cc_count_6_site_id") = iSiteId;
		if(sInnerUrl&&sInnerUrl!="undefined") doAlert("",sInnerUrl);
		return true;
	}

	function chkAdmin(){
		if(Session("cc_count_6_admin")==AdminPwd) return true;
		doAlert("","../admin/login.asp");return false;
	}

	function doLoginAdmin(pwd){
		Session("cc_count_6_admin")=pwd;
	}
	
	function formatMaxValue(x){
		var iBase = Math.pow(10,(x+'').length-1);
		if(iBase<10) iBase=10;
		return iBase-(x%iBase)+x;
	}

	function isPostBack(){
		if(Request.ServerVariables("REQUEST_METHOD")+""=="POST")
			return true;
		else
			return false;
	}

	function doAlert(s,u){
		Response.write("<"+"script language='JavaScript'>\n");
		if(s) Response.Write("alert('"+s.replace("'","\\'")+"');\n");
		if(u) Response.Write("location.href='"+u.replace("'","\\'")+"';\n");
		Response.write("</script"+">\n");
		if(u) Response.End();
	}

	function Str4Sql(s){
		return s.replace(/\'/gm,"''");
	}

	function getMonthName(n){
		var a = new Array("一","二","三","四","五","六","七","八","九","十","十一","十二");
		return a[n];
	}
	
    function getHttpHtml(sUrl){
		try{
			var sGetHtml;
			var oXmlHttp = Server.CreateObject("Msxml2.XmlHttp");
			oXmlHttp.open("GET",sUrl,false)
			oXmlHttp.send();
			sGetHtml = oXmlHttp.responseBody;
			oXmlHttp = null;
			return Bytes2bStr(sGetHtml);
		}catch(e){return '';}
    }
	
    function Bytes2bStr(vin){
        var StringReturn;
        var BytesStream = new ActiveXObject("ADODB.Stream");
        with(BytesStream){
			Type = 2;	//adTypeText
			Open();
			WriteText(vin);
			Position = 0;
			Charset = "GB2312";
			Position = 2;
			StringReturn = ReadText;
			Close();
        }
        BytesStream = null;
        return StringReturn;
    }
	
    function read(n){
		try{
			var sContent;
			var oSt = Server.CreateObject("ADODB.Stream");
			oSt.Charset = "GB2312";
			oSt.Open();
			oSt.LoadFromFile(Server.MapPath(n));
			sContent = oSt.ReadText();
			oSt = null;
			return sContent;
		}catch(e){return '';}
    }

	function write(n, s){
		try{
			var fso = Server.CreateObject("Scripting.FileSystemObject");
			var f1 = fso.CreateTextFile(Server.MapPath(n), true);
			f1.write(s);
			fso=f1=null;
			return 0;
		}catch(e){return 1;}
	}
</script>
<%
	var oQueryDate =  Request.QueryString('QueryDate');
	if(oQueryDate.Count==1){
		var aQueryDate = oQueryDate(1).split('-');
		if(aQueryDate.length==3) var sQueryDate = new Date(parseInt(aQueryDate[0]),parseInt(aQueryDate[1])-1,parseInt(aQueryDate[2]));
		else sQueryDate = new Date();
		if(isNaN(sQueryDate)) sQueryDate = new Date();
	}
	var iSiteId = 0;
	var dToday = oQueryDate.Count?sQueryDate:new Date();
	var dYesterday = new Date(dToday - (1000*60*60*24));
	var iSecOfMonth=(new Date(dToday.getYear(),dToday.getMonth()+1,1) - new Date(dToday.getYear(),dToday.getMonth(),1));
	var iDayOfMonth = parseInt(iSecOfMonth/1000/60/60/24);
	var dLastMonth = new Date(dToday - iSecOfMonth + 1000*60*60*24);
	var dLastWeek = new Date(dToday - (1000*60*60*24)*6);
	var sToday = formatDateTime(dToday,2);
	var sYesterday = formatDateTime(dYesterday,2);
	var sLastMonth = formatDateTime(dLastMonth,2);
	var sLastWeek = formatDateTime(dLastWeek,2);
	var iNowHour = dToday.getHours();
	var iDayOfWeek = 7;
	var iMonthOfYear = 12;
	
	ProjectName			=	"COCOON Counter 6 professional"
	Version				=	"1905"
	CodingInfo			=	"Coding  by  Sunrise_Chen."
	WebTitle			=	ProjectName + " [v:" + Version+"]"
	SysInfo				=	"Copyright(r) COCOON Studio 2002-2003."
	DeveloperInfo		=	"Sunrise_Chen"

	Date.prototype.getCnWeekName = function(){
		var a = new Array("日","一","二","三","四","五","六");
		return a[this.getDay()];
	}
%>
