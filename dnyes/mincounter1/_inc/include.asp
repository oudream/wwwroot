<script language="JScript" runat="server">
	//COCOON Counter 6 professional v:1900 (2003-12-19) by Sunrise Chen.
	function cc_app_counter_6(oConn){
		function formatDateTime(d,t){
			function formatNumber(n){return (n+'').length==1?'0'+n:n;}
			var sY=d.getFullYear();sM=d.getMonth()+1,sD=d.getDate();
			var sH=formatNumber(d.getHours()),sN=formatNumber(d.getMinutes());sS=formatNumber(d.getSeconds());
			var sDate = sY+'-'+sM+'-'+sD, sTime = sH+':'+sN+':'+sS;
			switch(t){
				case 0: return sDate+' '+sTime;
				case 1: case 2: return sDate;
				case 3:	case 4: return sTime;
			}
		}

		var oRs = new ActiveXObject('ADODB.Recordset');
		oRs.ActiveConnection = oConn;
		oRs.CursorType = 1;	//adOpenKeyset
		oRs.LockType = 3;	//adLockOptimistic
		
		var dtVisitTime = new Date((new Date()).valueOf()+TimeZoneOffset*3600000);
		var sTimeout = formatDateTime(new Date(dtVisitTime.valueOf()-1000*60*IpTimeExpire),0);
		var sVisitTime = formatDateTime(dtVisitTime,0);
		var sVisitDate = formatDateTime(dtVisitTime,1);
		var sHour = dtVisitTime.getHours();
		this.siteId = 0;
		
		this.setSite = function setSite(siteUid){	//初始化站点
			var sVarName = "cc_6_site_uid_"+siteUid ;
			var sSiteId = Application.Contents(sVarName);
			if(!isNaN(sSiteId)){ this.siteId = sSiteId; return; }
			sSqlString = "select siteid from t_Site where uid='"+siteUid+"'";
			var oRs = oConn.execute( sSqlString );
			if(oRs.EOF){
				oRs = null; oConn.Close(); oConn = null;
				Response.Write('/*站点未找到*/');
				Response.End();
			}
			this.siteId = oRs.Fields.Item(0).Value;
			Application.Contents(sVarName) = this.siteId;
			oRs = null;
		}
		
		this.isVisited = function isVisited(sIp){	//判断是否在规定时间内刷新 v6:1000
			var sVarName = "cc_6_user_isVisit_site_"+this.siteId ;
			if(!Session.Contents(sVarName)){
				Session.Contents(sVarName) = dtVisitTime;
				var sSqlString = "select 0 from t_Visit where siteid="+this.siteId+" and ip='"+sIp+"' and CreateTime>#"+sTimeout+"#";
				if(oConn.execute(sSqlString).EOF) return false;
				return true;
			}else{
				return true;
			}
		}
		
		this.addPageCount = function addPageCount(){	//访问统计 v6:beta2+ updated ; 1650 redesigned! ; 1700 fixed
			var sVarName = "cc_6_site_"+this.siteId+"_visited"
			if(Application.Contents(sVarName)!=sVisitDate){
				Application.Contents(sVarName)=sVisitDate;
				with(oRs){
					Source = "select siteid,s_"+sHour+" from t_Count where siteid="+this.siteId+" and CountDate=#"+sVisitDate+"#";
					Open(); if(EOF){AddNew([0,1],[this.siteId,1]);Update();} Close();
				}
				sSqlString = "delete from t_Visit where siteId="+this.siteId+" and CreateTime<#"+sVisitDate+"#";
				oConn.execute( sSqlString );
			}else{
				sSqlString = "update t_Count set s_"+sHour+"=s_"+sHour+"+1 where siteid="+this.siteId+" and CountDate=#"+sVisitDate+"#";
				oConn.execute( sSqlString );
			}
		}
		
		this.addCount = function addCount(){	//时间统计 v6:beta2+ added
			sSqlString = "update t_Count set h_"+sHour+"=h_"+sHour+"+1 where siteid="+this.siteId+" and CountDate=#"+sVisitDate+"#";
			oConn.execute( sSqlString );
		}
		
		this.addPage = function addPage(sPageUrl,sPage){	//访问页统计
			with(oRs){
				Source = "select siteid,PageCount,page,Url,CreateTime from t_Page where siteid="+this.siteId+" and Page='"+sPage+"'";
				Open(); if(EOF){AddNew([0,1,2,3],[this.siteId,1,sPage,sPageUrl]);}else{++Fields.Item(1).Value; Fields.Item(4).Value=sVisitTime}
				Update(); Close();
			}
		}
		
		this.addReferer = function addReferer(sRefUrl,sRefSite){	//引用页统计
			with(oRs){
				Source = "select siteid,RefPageCount,RefSite,RefPage,CreateTime from t_RefPage where siteid="+this.siteId+" and RefSite='"+sRefSite+"'";
				Open();
				if(EOF){
					AddNew([0,1,2,3],[this.siteId,1,sRefSite,sRefUrl]);
				}else{++Fields.Item(1).Value;
					Fields.Item(3).Value=sRefUrl;
					Fields.Item(4).Value=sVisitTime;
				}
				Update(); Close();
			}
		}
		
		this.addRegion = function addRegion(sRegion,sIp){	//地区统计
			with(oRs){
				Source = "select siteid,RegionCount,Region,RegionIp,CreateTime from t_Region where siteid="+this.siteId+" and Region='"+sRegion+"'";
				Open(); if(EOF){AddNew([0,1,2,3],[this.siteId,1,sRegion,sIp]);}else{++Fields.Item(1).Value; Fields.Item(3).Value=sIp;Fields.Item(4).Value=sVisitTime;}
				Update(); Close();
			}
		}
		
		this.addLinkType = function addLinkType(sLinkType){		//连接方式统计
			with(oRs){
				Source = "select siteid,LinkCount,Link from t_Link where siteid="+this.siteId+" and Link='"+sLinkType+"'";
				Open(); if(EOF) AddNew([0,1,2],[this.siteId,1,sLinkType]); else ++Fields.Item(1).Value;
				Update(); Close();
			}
		}
		
		this.addKeyword = function addKeyword(sKeyword,sRefUrl){	//关键字统计
			with(oRs){
				Source = "select siteid,KeywordCount,Keyword,RefPage,CreateTime from t_Keyword where siteid="+this.siteId+" and Keyword='"+sKeyword+"'";
				Open(); if(EOF){AddNew([0,1,2,3],[this.siteId,1,sKeyword,sRefUrl]);}else{++Fields.Item(1).Value; Fields.Item(4).Value=sVisitTime}
				Update(); Close();
			}
		}
		
		this.addClientInfo = function addClientInfo(sOS,sBrowser,sScreenSize,sColorDepth,sLanguage,sTimeZone){	//客户端信息统计
			with(oRs){
				Source = "select siteid,ClientCount,OS,Browser,ScreenSize,ColorDeep,SystemLanguage,TimeZone from t_Client " 
				+ "where siteid="+this.siteId+" "
				+ "and OS+Browser+ScreenSize+ColorDeep+SystemLanguage+TimeZone="
				+ "'"+sOS+sBrowser+sScreenSize+sColorDepth+sLanguage+sTimeZone+"'";
				Open(); if(EOF) AddNew([0,1,2,3,4,5,6,7],[this.siteId,1,sOS,sBrowser,sScreenSize,sColorDepth,sLanguage,sTimeZone]); else ++Fields.Item(1).Value;
				Update(); Close();
			}
		}
		
		this.addNewVisit = function addNewVisit(sPageUrl,sRefUrl,sIp,sRegion,sLinkType,sOS,sBrowser,sScreenSize,sColorDepth,sLanguage,sTimeZone){
			sSqlString = "insert into t_Visit(siteId,PageUrl,RefUrl,ip,region,linkType,ClientInfo)"
				+ "values("+this.siteId+",'"+sPageUrl+"','"+sRefUrl+"','"+sIp+"','"+sRegion+"','"+sLinkType+"','"+sOS+"\n"+sBrowser+"\n"+sScreenSize+"\n"+sColorDepth+"\n"+sLanguage+"\n"+sTimeZone+"')";
			oConn.execute( sSqlString );
		}
	}

	//类: 客户端信息
	function cc_class_ClientInfo(oConn){
		this.ip = '';
		this.region = '';
		this.linkType = '';
		this.browser = '';
		this.os = '';
		this.screenSize = '';
		this.colorDepth = '';
		this.systemLanguage = '';
		this.timeZone = '';
		this.UserAgentString = '';
		this.pageUrl = '';
		this.page = '';
		this.referrerPage = '';
		this.referrerSite = '';
		this.keyword = '';
		
		this.setUserAgent = function setUserAgent(sUserAgent){
			with(this){
				UserAgentString = sUserAgent;
				os = getOsInfo(sUserAgent);
				browser = getBrowserInfo(sUserAgent);
			}
		}
		
		this.setUserIp = function setUserIp(sUserIp){
			this.ip = sUserIp
			var sVarName = "cc_6_userIpSet";
			var sCookiesIp = Request.Cookies(sVarName)(sUserIp);
			if(sCookiesIp){
				var aCookiesIp = (sCookiesIp+'###').split('###');
				this.region = aCookiesIp[0];
				this.linkType = aCookiesIp[1];
				return;
			}
			var iIpNum = ipToNumber(sUserIp);
			//Call SQLServer procedure: getRegionByIp()
			var sSqlString = 'select top 1 country,city from t_ip_package where ip1<='+iIpNum+' and ip2>='+iIpNum+'';
			var oRs = oConn.execute( sSqlString );
			if(!oRs.EOF){
				this.region = oRs.Fields.Item(0).Value;
				this.linkType = getLink(oRs.Fields.Item(1).Value);
				Response.Cookies(sVarName)(sUserIp) = this.region + "###" + this.linkType;
				Response.Cookies(sVarName).Expires = formatDateTime(new Date((new Date()).valueOf()+7776000000),1); //1000*60*60*24*90,90天
			}
			oRs = null;
		}
		
		this.setPage = function setPage(sPageUrl){
			this.pageUrl = sPageUrl;
			this.page = sPageUrl.replace(/[\?#].*/,'');
		}
		
		this.setRefererPage = function setRefererPage(sRefUrl){
			if(!sRefUrl) return;
			this.referrerPage = sRefUrl;
			this.referrerSite = sRefUrl.replace(/(http:\/\/.+?)(\/.*)/gi,'$1');
			this.keyword = getSearchKeyword(sRefUrl);
		}
		
		this.setScreenSize = function setScreenSize(sScreenSize){	//屏幕大小 v6:1000
			if(isNaN(sScreenSize.replace('_','.'))) return;
			this.screenSize = sScreenSize;
		}
		
		this.setColorDepth = function setColorDepth(sColorDepth){	//颜色深度 v6:1000
			if(isNaN(sColorDepth)) return;
			this.colorDepth = sColorDepth;
		}
		
		this.setLanguage = function setLanguage(sSystemLanguage){	//语种 v6:1003
			this.systemLanguage = sSystemLanguage;
		}
		
		this.setTimeZone = function setTimeZone(sTimeZone){		//时区 v6:1007
			this.timeZone = sTimeZone;
		}
		
		function getLink(s){	//上网方式 v6:1000
			if(!s) return '';
			var p = "(电信|专线|长城|有线通|移动|广电|铁通|网通|联通|吉通|电通|ADSL|CHINANET|东方网景|网吧|大学|LAN|GPRS|拨号)";
			var r = new RegExp(p,'gim');
			var a = s.match(r);
			return a ? a[a.length-1] : '';
		}
		
		function getBrowserInfo(sUserAgent){	//浏览器 v6:1000
			var r = /(Netscape|Opera|NetCaptor|MSN |MSIE|MyIE|Mozilla)[\s\/]{0,1}\d{0,}\.{0,1}\d*/gim;
			var a = sUserAgent.match(r);
			return a ? a[a.length-1].replace(/\//g,' ') : '';
		}
		
		function getOsInfo(sUserAgent){		//操作系统 v6:1000
			var r = /(Windows|Mac_|Mac |unix|Linux|SunOS|BSD)[^;\(\)]+/gim;
			var a = sUserAgent.match(r);
			return a ? a[a.length-1] : '';
		}
		
		function getSearchKeyword(sRefererUrl){		//搜索关键字 v6:1000
			var p = "("
				+"google.+?q=([^&]*)"+"|sina.+?word=([^&]*)"
				+"|sohu.+?word=([^&]*)"+"|163.+?q=([^&]*)"
				+"|yahoo.+?p=([^&]*)"+"|baidu.+?word=([^&]*)"
				+"|openfind.+?q=([^&]*)"+"|lycos.+?query=([^&]*)"
				+"|aol.+?query=([^&]*)"+"|onseek.+?keyword=([^&]*)"
				+"|3721\.com.+?name=([^&]*)"+"|search\.tom.+?word=([^&]*)"
				+")";
			var r = new RegExp(p,'gim');
			var a = r.exec(sRefererUrl);
			if(a) for(var i=a.length-1;i>=0;--i) if(a[i]) try{return decodeURIComponent(a[i]);}catch(e){return a[i];}
			return '';
		}
		
		function ipToNumber(sIp){	//计算IP值 v4 in vbscript
			var a = sIp.split('.');
			var iNumber = 0;
			if(a.length!=4) return iNumber;
			for(var i=0;i<4;++i) iNumber += parseInt(a[i]) * Math.pow(256,3-i);
			return iNumber-1;
		}
	}
</script>

<!--#include file="Common.asp"-->