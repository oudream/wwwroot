<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	oConn.Open()
		try{
			sSqlString = "ALTER TABLE t_Client ADD COLUMN"
				+ " TimeZone varchar(4) NULL";
			oConn.Execute( sSqlString )
			Response.write("t_Client 更新成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。")
			Response.write(e.description)
		}
		Response.write("<br>\n")

		try{
			sSqlString = "DROP VIEW QueryMaxCount"
			oConn.Execute( sSqlString );
			Response.write("QueryMaxCount 删除成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。")
			Response.write(e.description)
		}
		Response.write("<br>\n");
		
		try{
			sSqlString = "DROP VIEW QueryMonthCount"
			oConn.Execute( sSqlString );
			Response.write("QueryMonthCount 删除成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。")
			Response.write(e.description)
		}
		Response.write("<br>\n");
		
		try{
			sSqlString = "DROP VIEW QueryMonthPageView"
			oConn.Execute( sSqlString );
			Response.write("QueryMonthPageView 删除成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。")
			Response.write(e.description)
		}
		Response.write("<br>\n");
		
		try{
			sSqlString = "DROP VIEW QueryPageCount"
			oConn.Execute( sSqlString );
			Response.write("QueryPageCount 删除成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。")
			Response.write(e.description)
		}
		Response.write("<br>\n");
		
		try{
			sSqlString = "DROP VIEW QueryTodayCount"
			oConn.Execute( sSqlString );
			Response.write("QueryTodayCount 删除成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。")
			Response.write(e.description)
		}
		Response.write("<br>\n");
		
		try{
			sSqlString = "DROP VIEW QueryTodayPageView"
			oConn.Execute( sSqlString );
			Response.write("QueryTodayPageView 删除成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。")
			Response.write(e.description)
		}
		Response.write("<br>\n");
		
		try{
			sSqlString = "DROP VIEW QueryTotalCount"
			oConn.Execute( sSqlString );
			Response.write("QueryTotalCount 删除成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。")
			Response.write(e.description)
		}
		Response.write("<br>\n");
		
		try{
			sSqlString = "DROP VIEW QueryYearCount"
			oConn.Execute( sSqlString );
			Response.write("QueryYearCount 删除成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。")
			Response.write(e.description)
		}
		Response.write("<br>\n");
		
		try{
			sSqlString = "DROP VIEW QueryYearPageView"
			oConn.Execute( sSqlString );
			Response.write("QueryYearPageView 删除成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。")
			Response.write(e.description)
		}
		Response.write("<br>\n");
		
		try{
			sSqlString = "DROP VIEW QueryYesterdayCount"
			oConn.Execute( sSqlString );
			Response.write("QueryYesterdayCount 删除成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。")
			Response.write(e.description)
		}
		Response.write("<br>\n");
		
		try{
			sSqlString = "DROP VIEW QueryYesterdayPageView"
			oConn.Execute( sSqlString );
			Response.write("QueryYesterdayPageView 删除成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。")
			Response.write(e.description)
		}
		Response.write("<br>\n");
		
		try{
			sSqlString = "DROP VIEW v_SiteCount"
			oConn.Execute( sSqlString );
			Response.write("v_SiteCount 删除成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。")
			Response.write(e.description)
		}
		Response.write("<br>\n");
		
		////////////////////////////
		Response.write("<br>\n");
		////////////////////////////
		
		try{
			sSqlString = "CREATE VIEW QueryMaxCount AS "
				+"SELECT a.SiteId, a.HourCount AS MaxCount, a.CountDate AS MaxDate FROM v_Count AS a, [SELECT SiteId, Max(HourCount) AS MaxCount FROM v_Count GROUP BY SiteId]. AS b WHERE a.SiteId=b.SiteId and a.HourCount=b.MaxCount;";
			oConn.Execute( sSqlString )
			Response.write("QueryMaxCount 创建成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。");
			Response.write(e.description);
		}
		Response.write("<br>\n")

		try{
			sSqlString = "CREATE VIEW QueryMonthCount AS "
				+"SELECT SiteId, Sum(HourCount) AS MonthCount, Sum(v_Count.PageViewCount) AS MonthPageView FROM v_Count WHERE CountDate>=DateSerial(Year(Date()),Month(Date()),1) GROUP BY SiteId;";
			oConn.Execute( sSqlString )
			Response.write("QueryMonthCount 创建成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。");
			Response.write(e.description);
		}
		Response.write("<br>\n")

		try{
			sSqlString = "CREATE VIEW QueryTodayCount AS "
				+"SELECT SiteId, Sum(HourCount) AS TodayCount, Sum(v_Count.PageViewCount) AS TodayPaegView FROM v_Count WHERE CountDate=Date() GROUP BY SiteId;";
			oConn.Execute( sSqlString )
			Response.write("QueryTodayCount 创建成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。");
			Response.write(e.description);
		}
		Response.write("<br>\n")

		try{
			sSqlString = "CREATE VIEW QueryTotalCount AS "
				+"SELECT SiteId, Sum(HourCount) AS TotalCount, Sum(v_Count.PageViewCount) AS TotalPageCount FROM v_Count GROUP BY SiteId;";
			oConn.Execute( sSqlString )
			Response.write("QueryTotalCount 创建成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。");
			Response.write(e.description);
		}
		Response.write("<br>\n")

		try{
			sSqlString = "CREATE VIEW QueryYearCount AS "
				+"SELECT SiteId, Sum(HourCount) AS YearCount, Sum(v_Count.PageViewCount) AS YearPageView FROM v_Count WHERE CountDate>=DateSerial(Year(Date()),1,1) GROUP BY SiteId;";
			oConn.Execute( sSqlString )
			Response.write("QueryYearCount 创建成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。");
			Response.write(e.description);
		}
		Response.write("<br>\n")

		try{
			sSqlString = "CREATE VIEW QueryYesterdayCount AS "
				+"SELECT SiteId, Sum(HourCount) AS YesterdayCount, Sum(v_Count.PageViewCount) AS YesterdayPageView FROM v_Count WHERE CountDate=Date()-1 GROUP BY SiteId;";
			oConn.Execute( sSqlString )
			Response.write("QueryYesterdayCount 创建成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。");
			Response.write(e.description);
		}
		Response.write("<br>\n")

		try{
			sSqlString = "CREATE VIEW v_SiteCount AS "
				+"SELECT QueryDate.CountFromDate, QueryDate.CountToDate, QueryDate.CountDays, QueryTotalCount.TotalCount, QueryTotalCount.TotalPageCount, QueryMaxCount.MaxCount, QueryMaxCount.MaxDate, QueryYesterdayCount.YesterdayCount, QueryTodayCount.TodayCount, QueryMonthCount.MonthCount, QueryYearCount.YearCount, QueryYesterdayCount.YesterdayPageView, QueryTodayCount.TodayPaegView, QueryMonthCount.MonthPageView, QueryYearCount.YearPageView, t_Site.* FROM ((((((t_Site LEFT JOIN QueryDate ON t_Site.SiteId = QueryDate.SiteId) LEFT JOIN QueryTotalCount ON t_Site.SiteId = QueryTotalCount.SiteId) LEFT JOIN QueryMaxCount ON t_Site.SiteId = QueryMaxCount.SiteId) LEFT JOIN QueryYearCount ON t_Site.SiteId = QueryYearCount.SiteId) LEFT JOIN QueryTodayCount ON t_Site.SiteId = QueryTodayCount.SiteId) LEFT JOIN QueryMonthCount ON t_Site.SiteId = QueryMonthCount.SiteId) LEFT JOIN QueryYesterdayCount ON t_Site.SiteId = QueryYesterdayCount.SiteId;";
			oConn.Execute( sSqlString )
			Response.write("v_SiteCount 创建成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。");
			Response.write(e.description);
		}
		Response.write("<br>\n");

	sSqlString = "select siteid,max(CountDate) from t_Count group by siteid"
	var oRs = oConn.execute( sSqlString );
	while(!oRs.EOF){
		Application.Contents("cc_6_site_"+oRs.fields.item(0).value+"_visited")=formatDateTime(new Date(oRs.fields.item(1).value),1);
		
		Response.write("<br>");
		Response.write("cc_6_site_"+oRs.fields.item(0).value+"_visited");
		Response.write(" = #");
		Response.write(Application.Contents("cc_6_site_"+oRs.fields.item(0).value+"_visited"));
		Response.write("#<p>");
		oRs.moveNext();
	}

	oConn.Close()
%>


1650 更新完成。
