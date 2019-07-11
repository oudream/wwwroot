<!--#include file=conn.asp -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="md5.asp"-->
<!--#include file="ChkURL.asp"-->
<!--#include file="ChkManage.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	action=request("action")


	'if action="UpdateReview" then
		'对 NEWS 表中的相关评论标志[titlesize]进行更新
		conn.execute("update "& db_News_Table &" Set titlesize=0 ")
		
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_Review_Table &" where NewsID>0 order by NewsID" 
		rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "<DIV align=center><font color=red>本系统还没有文章评论！</font></DIV><BR>"
		else
			do while not rs.eof
				NewsIDS=rs("NewsID")
		
				set rst=server.createobject("adodb.recordset")
				sqlt="select * from "& db_News_Table &" where NewsId="&NewsIDS
				rst.open sqlt,conn,1,3
				if rst.eof and rst.bof then
					rst.close
					set rst=nothing
					response.write "<DIV align=center><font color=red>---无法找到对应评论的文章---</font></DIV><BR>"
				else
					rst("titlesize")=rst("titlesize")+1
					rst.update
					rst.close
					set rst=nothing
					response.write "<DIV align=center>" &NewsIDS& "号文章评论标记更新成功。</DIV><BR>"
				end if
				rs.movenext
			loop
			rs.close
			set rs=nothing
			response.write "<DIV align=center>文章评论标记更新完毕！</DIV><BR>"
		end if
		
		'更新表结构
		 conn.execute("alter table dep drop deptype ")
         conn.execute("alter table dep add deptype varchar(50)")
         response.write "<DIV align=center>修改字段[dep]表[deptype]属性完毕！</DIV><BR>"

		 conn.execute("alter table admin drop deptype ")
         conn.execute("alter table admin add deptype varchar(50)")
         response.write "<DIV align=center>修改字段[admin]表[deptype]属性完毕！</DIV><BR>"

		 conn.execute("alter table FT_User drop deptype ")
         conn.execute("alter table FT_User add deptype varchar(50)")
         response.write "<DIV align=center>修改字段[FT_User]表[deptype]属性完毕！</DIV><BR>"

         conn.execute("alter table news add backtype int")
         response.write "<DIV align=center>添加[news]段落格式字段[backtype]完毕！</DIV><BR>"

        '更新表数据	

		conn.execute("update "& db_News_Table &" Set titleface='无' where titleface='' or titleface='0' or (titleface is null)")
		response.write "<DIV align=center>文章链接地址标记更新完毕！</DIV><BR>"

		conn.execute("update "& db_Review_Table &" Set passed=1 where NewsID<1 or (NewsID is null)")
		response.write "<DIV align=center>留言默认审核数据标记更新完毕！</DIV><BR>"

		conn.execute("update "& db_News_Table &" Set backtype=0")
		response.write "<DIV align=center>段落格式标记更新完毕！</DIV><BR>"

		conn.execute("update "& db_System_Table &" Set M_MAIN=0")
		response.write "<DIV align=center>签收标记更新完毕！</DIV><BR>"

		response.write "<DIV align=center>全部更新完成！</DIV><BR>"
		response.end
		conn.close
		set conn=nothing
	'else 
	'	response.write "<DIV align=center><font color=red>请选择一个参数！</font></DIV>"
	'	response.end
	'end if
end if%>