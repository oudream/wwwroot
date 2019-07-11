<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	oConn.Open()
		try{
			sSqlString = "DROP VIEW v_Count";
			oConn.Execute( sSqlString )
			Response.write("v_Count 删除成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经从升过级了。")
			Response.write(e.description)
		}
		Response.write("<br>\n")
		
		try{
			sSqlString = "CREATE VIEW v_Count AS "
				+ "SELECT t_Count.SiteId, t_Count.CountDate," 
				+ "(h_0+h_1+h_2+h_3+h_4+h_5+h_6+h_7+h_8+h_9+h_10+h_11+h_12+h_13+h_14+h_15+h_16+h_17+h_18+h_19+h_20+h_21+h_22+h_23)"
				+ " AS HourCount, " 
				+ "(s_0+s_1+s_2+s_3+s_4+s_5+s_6+s_7+s_8+s_9+s_10+s_11+s_12+s_13+s_14+s_15+s_16+s_17+s_18+s_19+s_20+s_21+s_22+s_23)" 
				+ " AS PageViewCount FROM t_Count;";
			oConn.Execute( sSqlString )
			Response.write("v_Count 创建成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。");
			Response.write(e.description);
		}
		Response.write("<br>\n")

	oConn.Close()
%>


1850 更新完成。
