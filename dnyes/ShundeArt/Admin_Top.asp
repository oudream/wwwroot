<table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr>
<td background="IMAGES/top1.gif">
	<table border="0" cellpadding="6" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
        <tr> 
          <td background="IMAGES/top1.gif">&nbsp;</td>
        </tr>
        <tr> 
          <td background="IMAGES/top1.gif" class="td001"> <font color="#FFFFFF">
		 �û���<b><%=Request.cookies(Forcast_SN)("UserName")%></b>
			<%if db_Birthday_Select="FeiTium" then		'�Ա��ֶ���ԭ�ֶ�%>
				<%=Request.cookies(Forcast_SN)("sex")%>
			<%else
				if db_Birthday_Select="Text" then		'�Ա��ֶ���BBS���ı��Ͱ���������
					if Request.cookies(Forcast_SN)("sex")=1 then%>
						��
					<%else
						if Request.cookies(Forcast_SN)("sex")=0 then%>
							Ů
						<%else%>
							����
						<%end if
					end if
				end if
			end if%>
					����Ȩ�ޣ�<%if Request.cookies(Forcast_SN)("KEY")="super" and Request.cookies(Forcast_SN)("purview")="99999" then%><font color="#FFCC00">��������Ա</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="super" and Request.cookies(Forcast_SN)("purview")<>"99999" then%><font color="#FFCC00">ϵͳ����Ա</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="check" then%><font color="#FFCC00">�������Ա</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="selfreg" then%><font color="#FFCC00">ע���û�</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="smallmaster" then%><font color="#FFCC00">С�����Ա</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="bigmaster" then%><font color="#FFCC00">�������Ա</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="typemaster" then%><font color="#FFCC00">��������Ա</font><%end if%>
					���ĵȼ���<%if Request.cookies(Forcast_SN)("KEY")<>"selfreg" then%><font color="#FFCC00">�ڲ���Ա</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="selfreg" and Request.cookies(Forcast_SN)("reglevel")="1" then%><font color="red">��ͨ</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="selfreg" and Request.cookies(Forcast_SN)("reglevel")="2" then%><font color="red">�߼�</font><%end if%>
						<%if Request.cookies(Forcast_SN)("KEY")="selfreg" and Request.cookies(Forcast_SN)("reglevel")="3" then%><font color="red">�ؼ�</font><%end if%>
						��¼������<%=Request.cookies(Forcast_SN)("UserLoginTimes")%>
		  </font>
        </td>
        </tr>
      </table>
</td>
</tr>
<tr>
	<td bgcolor="#7C7CB5">
      <div align="right">
	  <A href="Admin_Exit.asp"><FONT color=#ffffff>�˳���վ</FONT></A>&nbsp;&nbsp;
	  <A onclick="checkclick('���Ƿ�Ҫ�򿪱�վ��ҳ��')" href="index.asp" target=_blank><FONT color=#ffffff>��վ��ҳ</FONT></A></div></td>
</tr>
</table>
