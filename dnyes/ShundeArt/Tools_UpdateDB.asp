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
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	action=request("action")


	'if action="UpdateReview" then
		'�� NEWS ���е�������۱�־[titlesize]���и���
		conn.execute("update "& db_News_Table &" Set titlesize=0 ")
		
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_Review_Table &" where NewsID>0 order by NewsID" 
		rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "<DIV align=center><font color=red>��ϵͳ��û���������ۣ�</font></DIV><BR>"
		else
			do while not rs.eof
				NewsIDS=rs("NewsID")
		
				set rst=server.createobject("adodb.recordset")
				sqlt="select * from "& db_News_Table &" where NewsId="&NewsIDS
				rst.open sqlt,conn,1,3
				if rst.eof and rst.bof then
					rst.close
					set rst=nothing
					response.write "<DIV align=center><font color=red>---�޷��ҵ���Ӧ���۵�����---</font></DIV><BR>"
				else
					rst("titlesize")=rst("titlesize")+1
					rst.update
					rst.close
					set rst=nothing
					response.write "<DIV align=center>" &NewsIDS& "���������۱�Ǹ��³ɹ���</DIV><BR>"
				end if
				rs.movenext
			loop
			rs.close
			set rs=nothing
			response.write "<DIV align=center>�������۱�Ǹ�����ϣ�</DIV><BR>"
		end if
		
		'���±�ṹ
		 conn.execute("alter table dep drop deptype ")
         conn.execute("alter table dep add deptype varchar(50)")
         response.write "<DIV align=center>�޸��ֶ�[dep]��[deptype]������ϣ�</DIV><BR>"

		 conn.execute("alter table admin drop deptype ")
         conn.execute("alter table admin add deptype varchar(50)")
         response.write "<DIV align=center>�޸��ֶ�[admin]��[deptype]������ϣ�</DIV><BR>"

		 conn.execute("alter table FT_User drop deptype ")
         conn.execute("alter table FT_User add deptype varchar(50)")
         response.write "<DIV align=center>�޸��ֶ�[FT_User]��[deptype]������ϣ�</DIV><BR>"

         conn.execute("alter table news add backtype int")
         response.write "<DIV align=center>���[news]�����ʽ�ֶ�[backtype]��ϣ�</DIV><BR>"

        '���±�����	

		conn.execute("update "& db_News_Table &" Set titleface='��' where titleface='' or titleface='0' or (titleface is null)")
		response.write "<DIV align=center>�������ӵ�ַ��Ǹ�����ϣ�</DIV><BR>"

		conn.execute("update "& db_Review_Table &" Set passed=1 where NewsID<1 or (NewsID is null)")
		response.write "<DIV align=center>����Ĭ��������ݱ�Ǹ�����ϣ�</DIV><BR>"

		conn.execute("update "& db_News_Table &" Set backtype=0")
		response.write "<DIV align=center>�����ʽ��Ǹ�����ϣ�</DIV><BR>"

		conn.execute("update "& db_System_Table &" Set M_MAIN=0")
		response.write "<DIV align=center>ǩ�ձ�Ǹ�����ϣ�</DIV><BR>"

		response.write "<DIV align=center>ȫ��������ɣ�</DIV><BR>"
		response.end
		conn.close
		set conn=nothing
	'else 
	'	response.write "<DIV align=center><font color=red>��ѡ��һ��������</font></DIV>"
	'	response.end
	'end if
end if%>