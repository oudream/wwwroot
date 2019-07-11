<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->

<%msntime=now()   'if %>
<%response.cookies(Forcast_SN)("msntip")=now()%>
<LINK href=news.css rel=stylesheet>

<SCRIPT language=javascript>
window.onload = getMsg;
window.onresize = resizeDiv;
window.onerror = function(){}

var divTop,divLeft,divWidth,divHeight,docHeight,docWidth,objTimer,i = 0;
function getMsg()
{
    try{
    divTop = parseInt(document.getElementById("eMeng").style.top,10)
    divLeft = parseInt(document.getElementById("eMeng").style.left,10)
    divHeight = parseInt(document.getElementById("eMeng").offsetHeight,10)
    divWidth = parseInt(document.getElementById("eMeng").offsetWidth,10)
    docWidth = document.body.clientWidth;
    docHeight = document.body.clientHeight;
    document.getElementById("eMeng").style.top = parseInt(document.body.scrollTop,10) + docHeight + 10;//  divHeight
    document.getElementById("eMeng").style.left = parseInt(document.body.scrollLeft,10) + docWidth - divWidth
    document.getElementById("eMeng").style.visibility="visible"
    objTimer = window.setInterval("moveDiv()",1)
    }
    catch(e){}
}

function resizeDiv()
{
    i+=1
    if(i>1000) closeDiv()
    try{
    divHeight = parseInt(document.getElementById("eMeng").offsetHeight,10)
    divWidth = parseInt(document.getElementById("eMeng").offsetWidth,10)
    docWidth = document.body.clientWidth;
    docHeight = document.body.clientHeight;
    document.getElementById("eMeng").style.top = docHeight - divHeight + parseInt(document.body.scrollTop,10)
    document.getElementById("eMeng").style.left = docWidth - divWidth + parseInt(document.body.scrollLeft,10)
    }
    catch(e){}
}

function moveDiv()
{
    try
    {
    if(parseInt(document.getElementById("eMeng").style.top,10) <= (docHeight - divHeight + parseInt(document.body.scrollTop,10)))
    {
    window.clearInterval(objTimer)
    objTimer = window.setInterval("resizeDiv()",1)
    }
    divTop = parseInt(document.getElementById("eMeng").style.top,10)
    document.getElementById("eMeng").style.top = divTop - 1
    }
    catch(e){}
}
function closeDiv()
{
    document.getElementById('eMeng').style.visibility='hidden';
    if(objTimer) window.clearInterval(objTimer)
}
</SCRIPT>

<DIV id=eMeng style="BORDER-RIGHT: #455690 1px solid; BORDER-TOP: #a6b4cf 1px solid; Z-INDEX: 99999; LEFT: 0px; VISIBILITY: hidden; BORDER-LEFT: #a6b4cf 1px solid; WIDTH: 199px; BORDER-BOTTOM: #455690 1px solid; POSITION: absolute; TOP: 0px; HEIGHT: 97px; BACKGROUND-COLOR: #c9d3f3">
<TABLE oncontextmenu=return(false) style="BORDER-TOP: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid" cellSpacing=0 cellPadding=0 width="100%" bgColor=#cfdef4 border=0>
<TBODY>
	<TR>
		<TD style="FONT-SIZE: 12px; BACKGROUND-IMAGE: url(images/msgtopbg.gif); COLOR: #0f2c8c" width=30 height=23></TD>
		<TD style="PADDING-LEFT: 4px; FONT-WEIGHT: normal; FONT-SIZE: 12px; BACKGROUND-IMAGE: url(images/msgtopbg.gif); COLOR: #1f336b; PADDING-TOP: 4px" vAlign=center width="100%" height=23>【<%=jjgn%>】</TD>
		<TD style="PADDING-RIGHT: 2px; BACKGROUND-IMAGE: url(images/msgtopbg.gif); PADDING-TOP: 2px" vAlign=center align=right width=19 height=23><IMG title=关闭 style="CURSOR: hand" onclick=closeDiv() height=13 hspace=3 src="images/msgclose.gif" width=13></TD>
	</TR>
	<TR vAlign=top>
		<TD style="PADDING-RIGHT: 1px; BACKGROUND-IMAGE: url(images/msg_main.gif); PADDING-BOTTOM: 1px" colSpan=3 height=6>
			<DIV style="BORDER-RIGHT: #b9c9ef 1px solid; PADDING-RIGHT: 13px; BORDER-TOP: #728eb8 1px solid; PADDING-LEFT: 13px; FONT-SIZE: 12px; PADDING-BOTTOM: 13px; BORDER-LEFT: #728eb8 1px solid; WIDTH: 100%; COLOR: #1f336b; WORD-BREAK: break-all; PADDING-TOP: 1px; BORDER-BOTTOM: #b9c9ef 1px solid; HEIGHT: 100%" align=center>
				<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
				<TBODY>
					<TR>
						<TD height=8></TD>
					</TR>
					<TR>
						<TD height=34>
							<DIV align=center><IMG src="images/light.gif"></DIV>
						</TD>
					</TR>
					<TR>
						<TD height=28>
							<DIV style="FONT-SIZE: 12px">
								<DIV align=center>
									<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
									<TBODY>
										<TR>
											<TD height=17>
<%'if %>
<DIV align=center><FONT color=#0000cc><A href="" target=_blank>请您注册或登陆论坛</A></FONT></DIV>
											</TD>
										</TR>
										<TR>
											<TD height=17>
<DIV align=center><FONT color=#0000cc><A href="" target=_blank><%=request.cookies(Forcast_SN)("msntip")%></A></FONT></DIV>
											</TD>
										</TR>
										<TR>
											<TD height=20>
<DIV align=center><A href="" target=_blank><U><FONT color=#ff0000><%=(request.cookies(Forcast_SN)("msntip")-msntime)%>详情查看</FONT></U></A></DIV>
											</TD>
										</TR>
									</TBODY>
									</TABLE>
								</DIV>
							</DIV>
						</TD>
					</TR>
				</TBODY>
				</TABLE>
			</DIV>
		</TD>
	</TR>
</TBODY>
</TABLE>
</DIV>
