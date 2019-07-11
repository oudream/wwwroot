<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file=config.asp -->

<script language="Javascript">
function openScript(url, width, height){
var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}
function openem()
{
openScript('upload.asp',350,200);
}
</script>
<script language="javascript">
<!--
function isok(theform)
{
    if (theform.name.value=="")
    {
        alert("请填写广告名称！");
        theform.name.focus();
        return (false);
    }
    if (theform.url.value=="")
    {
        alert("请填写链接URL！");
        theform.url.focus();
        return (false);
    }
    return (true);
}
-->
</script>
<%
if request.querystring("job")="add" then



Call  addrk()
end if


%>

<!--#include file="top.asp"-->


              <table border=0 width=100% cellspacing=0 cellpadding=0 style="border-collapse: collapse" bordercolor="#111111">
            <tr> 
              <td align=center> 
                <%
if request.querystring("job")="suc" then
%>
                <font size="3" color=red><b>添加新广告条成功，请继续添加...</b></font> 
                <%
else
%>
                <font size="3" color=red><b>添加新广告条</b></font> 
                <%
end if
%>
     <hr color="#808080" size="1">         </td>
            </tr>
          </table>
              <table border=0 width=100% cellspacing=1 cellpadding=2>
				<form method="POST"  name="form"  action=addads.asp?job=add onSubmit="return isok(this)">

              <tr bgcolor=#ffffff> 
                <td width=85>广告名称</td>
                <td colspan="2"> 
                  <input type=text name=name size=30 maxlength=30>
                  &#12288;&#19981;&#36229;&#36807;15&#20010;&#20013;&#25991;&#25110;30&#20010;&#23383;&#27597;&#25968;&#23383;</td>
              </tr>
              <tr bgcolor=#ffffff> 
                <td width=85>链接URL</td>
                <td colspan="2"> 
                  <input type=text name=url size=40 value="http://">
                </td>
              </tr>
               <tr bgcolor=#ffffff> 
                <td width=85>&#31616;&#20171;/内容</td>
                <td width="200"> 
                  <textarea rows="5" name="intro" cols="38"></textarea></td>
                <td> <font color="#FF0000">提示：</font><br>
                  <font color="#808080">如果是嵌入代码请将代码内容填入此处 链接URL无效<br>
                  如果显示纯文本，则显示为广告名称<br>
                  只有GIF图片、Flash动画时URL填写有效</font></font>
                  
                  
                  </td>
              </tr>
              <tr bgcolor=#ffffff> 
                <td width=85>广告类型</td>
                <td colspan="2"> 
                  <input type="radio" name="xslei" value="gif">GIF图片 
                  <input type="radio" name="xslei" value="swf" checked><font siz=3 >Flash动画 </font>
                  <input type="radio" name="xslei" value="txt"><font siz=3 >纯文本 </font>    
                  <input type="radio" name="xslei" value="dai">嵌入代码
                </td>
              </tr>
              <tr bgcolor=#ffffff> 
                <td>&#22270;&#29255;URL</td>
                <td colspan="2"> 
                  <input type=text name=gif_url size=40 value="http://">
[<a href="JavaScript:openem()"><font color="blue">上传图片</font></a>(<font color="#FF0000">2M以内</font>)]<font color=red siz=3 > </font><font siz=3 >高度 </font> 
                  <input type=text name=hei size=3 maxlength="4"><font siz=3 >&nbsp; 
                  宽度 </font> 
                  <input type=text name=wid size=3 maxlength="4"><font color=red siz=3 > 
                  可以是百分比或空默认</font>
                </td>
              </tr>
              <tr bgcolor=#ffffff> 
                <td>所属广告位</td>
                <td colspan="2"> 
                
                <%adsconn.Open adsdata
                Call  Ggwxlxx(1) %>
                  
                  
                  　</td>
              </tr>
              <tr bgcolor=#ffffff> 
                <td>链接&#25171;&#24320;&#26041;&#24335;</td>
                <td colspan="2"> 
                  <select size=1 name=window>
                    <option value=0 selected>新窗口打开</option>
                    <option value=1>原窗口打开</option>
                  </select>
                </td>
              </tr>
              <tr bgcolor=#c0d0e0> 
                <td bgcolor="#DBDBDB" height="30">&#25773;&#25918;&#26465;&#20214;组</td>
                <td bgcolor="#DBDBDB" height="30" colspan="2">&#12288;在八个条件组中，任选其中一组&#65292;用于限制&#35813;&#24191;&#21578;&#33258;&#21160;&#36827;&#20837;&#20241;&#30496;&#29366;&#24577;的条件&#65292;以后&#21487;随时&#20462;&#25913;。</td>
              </tr>
              <tr bgcolor=#ffffff> 
                <td align=right><font color=red>（1）</font><input type=radio value=0 name=class checked></td>
                <td colspan="2"> 
                  <table border=0>
                    <tr> 
                      <td> 
                        　</td>
                      <td>&#26080;&#38480;&#21046;&#24490;&#29615;</td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td align=right bgcolor="#F5F5F5"><font color=red>（2）</font><input type=radio value=1 name=class></td>
                <td bgcolor=#ffffee colspan="2"> 
                  <table border=0>
                    <tr> 
                      <td> 
                        　</td>
                      <td>&#28857;&#20987;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=clicks1 size=8>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td align=right><font color=red>（3）</font><input type=radio value=2 name=class></td>
                <td bgcolor=#ffffff colspan="2"> 
                  <table border="0">
                    <tr> 
                      <td> 
                        　</td>
                      <td>&#26174;&#31034;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=shows2 size=8>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
               <td align=right bgcolor="#F5F5F5"><font color=red>（4）</font><input type=radio value=3 name=class></td>
                <td bgcolor=#ffffee colspan="2"> 
                  <table border="0">
                    <tr> 
                      <td> 
                        　</td>
                      <td>&#26174;&#31034;&#25130;&#27490;&#26399;&#20026;</td>
                      <td> 
                        <input type=text name=time3 size=20>
                        &#12288;<font color=#FF0000>&#26684;&#24335;&#20026;yyyy-mm-dd hh:mm:ss&#25110;&#32773;yyyy-mm-dd</font></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
               <td align=right><font color=red>（5）</font><input type=radio value=4 name=class></td>
                <td bgcolor=#ffffff colspan="2"> 
                  <table border="0">
                    <tr> 
                      <td> 
                        　</td>
                      <td>&#28857;&#20987;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=clicks4 size=8>
                      </td>
                    </tr>
                    <tr> 
                      <td align=center> </td>
                      <td>&#26174;&#31034;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=shows4 size=8>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
               <td align=right bgcolor="#F5F5F5"><font color=red>（6）</font><input type=radio value=5 name=class></td>

                <td bgcolor=#ffffee colspan="2"> 
                  <table border="0">
                    <tr> 
                      <td> 
                        　</td>
                      <td>&#28857;&#20987;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=clicks5 size=8>
                      </td>
                    </tr>
                    <tr> 
                      <td align=center> </td>
                      <td>&#26174;&#31034;&#25130;&#27490;&#26399;&#20026;</td>
                      <td> 
                        <input type=text name=time5 size=20>
                        &#12288;<font color=#FF0000>&#26684;&#24335;&#20026;yyyy-mm-dd hh:mm:ss&#25110;&#32773;yyyy-mm-dd</font></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td align=right><font color=red>（7）</font><input type=radio value=6 name=class></td>

                <td bgcolor=#ffffff colspan="2"> 
                  <table border="0">
                    <tr> 
                      <td> 
                        　</td>
                      <td>&#26174;&#31034;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=shows6 size=8>
                      </td>
                    </tr>
                    <tr> 
                      <td></td>
                      <td>&#26174;&#31034;&#25130;&#27490;&#26399;&#20026;</td>
                      <td> 
                        <input type=text name=time6 size=20>
                        &#12288;<font color=#FF0000>&#26684;&#24335;&#20026;yyyy-mm-dd hh:mm:ss&#25110;&#32773;yyyy-mm-dd</font></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
              <td align=right bgcolor="#F5F5F5"><font color=red>（8）</font><input type=radio value=7 name=class></td>

                <td bgcolor=#ffffee colspan="2"> 
                  <table border="0">
                    <tr> 
                      <td> 
                        　</td>
                      <td>&#28857;&#20987;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=clicks7 size=8>
                      </td>
                    </tr>
                    <tr> 
                      <td></td>
                      <td>&#26174;&#31034;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=shows7 size=8>
                      </td>
                    </tr>
                    <tr> 
                      <td></td>
                      <td>&#26174;&#31034;&#25130;&#27490;&#26399;&#20026;</td>
                      <td> 
                        <input type=text name=time7 size=20>
                        &#12288;<font color=#FF0000>&#26684;&#24335;&#20026;yyyy-mm-dd hh:mm:ss&#25110;&#32773;yyyy-mm-dd</font></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor=#ffffff> 
                <td colspan=3 align=center> 
                  <input type=submit value=&#25552;&#20132; name=B1>
                  <input type=reset value=&#37325;&#20889; name=B2>
                </td>
              </tr>
            </form>
          </table>
        
  <!--#include file="boot.asp"-->