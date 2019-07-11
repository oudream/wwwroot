<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file=config.asp -->
<script language=JavaScript>
<!--

function opw(url,name, width, height) { //v2.0
window.open(url,name,''+'width='+width+',height='+height+',scrollbars=yes'+'');
}
//-->
</script>

    <!--#include file="top.asp"-->
<br>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
    <tr>
      <td width="100%"><b>一、系统特点：</b><ol>
        <li>通过本系统可以设置并管理无数个广告位</li>
        <li>各广告位中可添加无数个循环播放的广告条</li>
        <li>
        广告位中的广告条已有7种显示方式,即&quot;页内嵌入循环&quot;、&quot;上下排列置入&quot;、&quot;左右排列置入&quot;、&quot;向上滚动置入&quot;、&quot;向左滚动置入&quot;、&quot;弹出多个窗口&quot;、&quot;循环弹出窗口&quot;，具体说明请参阅 
        <a href="addadw.asp#说明">广告位栏目中广告位显示方式说明</a></li>
        <li>广告条可以是GIF、SWF（Flash）或纯文本三种显示类型</li>
        <li>广告位上的广告条为循环播放，每次显示的是该广告位中等待时间最长、且处于正常状态的广告条</li>
        <li>可对任意广告条，随时执行暂停、激活、修改、删除等操作</li>
        <li>删除某一条广告时，与其相关的显示、点击记录也将随之删除</li>
        <li>轻松实现广告位的页面发布,具体参阅《<a href="#三">广告位发布说明</a>》</li>
        <li>多种广告播放条件控制广告播放状态，可设点击最高限制、显示最高限制、最后时间限制等</li>
        <li>完善的广告访问记录，可显示广告浏览者、点击者的IP地址</li>
        <li>当有大量广告条存在时，可通过多种条件查询广告条以对其进行操作</li>
      </ol>
      <p><b>二、使用说明：</b></p>
      <ol>
        <li>在 <font color="#FF0000">广 告 位</font> 一栏内可添加新广告位或修改、删除已有广告位标识，查询广告位ID</li>
        <li>在 <font color="#FF0000">添加广告 </font>一栏内可为某广告位添加一个新广告条</li>
        <li>在 <font color="#FF0000">正常广告 </font>
        一栏内显示当前所有处于正常播放状态的广告条，并可执行修改、删除、暂停、预览操作</li>
        <li>在 <font color="#FF0000">图片广告 </font>
        一栏内显示当前所有处于正常播放状态的非文本广告条，并可执行修改、删除、暂停、预览操作</li>
        <li>在 <font color="#FF0000">文本广告 </font>
        一栏内显示当前所有处于正常播放状态的纯文本广告条，并可执行修改、删除、暂停、预览操作</li>
        <li>在 <font color="#FF0000">点击排行 </font>内 
        按点击次数的不同顺序显示各广告条的点击次数，并可执行修改、删除、暂停、激活、预览操作</li>
        <li>在 <font color="#FF0000">显示排行 </font>内 
        按显示次数的不同顺序显示各广告条的显示次数，并可执行修改、删除、暂停、激活、预览操作</li>
        <li>在 <font color="#FF0000">暂停列表 </font>内 
        显示当前所有处于暂停播放状态的广告条，并可执行修改、删除、激活、预览操作</li>
        <li>在 <font color="#FF0000">失效列表 </font>内 
        显示当前所有已经失效的广告条，并可执行修改、删除、激活、预览操作</li>
        <li>在 <font color="#FF0000">广 告 位 </font>内 
        通过某广告位连接，可显示该广告位下的所有广告条，并可执行修改、删除、暂停、预览操作</li>
      </ol>
      <p><b><a name="三">三</a>、广告位发布说明：</b></p>
      <ol>
        <li>确定 <font color="#FF0000">实际页面中的预定广告位置</font> 应放置哪个 
        <font color="#FF0000">通过本系统设置的广告位</font> </li><br><br>
        <li>通过 <font color="#FF0000">广 告 位</font> 一栏，得到所需 <font color="#FF0000">
        广告位ID</font></li><br><br>
        <li>然后将下表的内容拷贝到预定广告位置，注意将其中的 <font color="#FF0000">广告位ID</font> 对应正确</li><br><br>
       

        <input type="text" name="T1" size="100" value="<script language=javascript src=<%=DqUrl()%>/ad.asp?i=广告位ID></script>"></li>
      </ol>

      <p><b>四、注意事项：</b></p>
      <ol>
        <li>每个广告位中的所有广告条显示图片宽度、高度应尽量保持一致，并应注意跟广告位预定的实际页面位置风格一致</li><br><br>
        <li>在实际页面预定的不同广告位中尽量放置使用本系统设置的不同广告位，这样可尽可能多地投放广告</li><br><br>
        <li>同一广告位中,文字广告条与图片广告条尽量不要混合使用</li>
      </ol>
      <p><font color="#FF0000"><b>备注：实际页面中的预定广告位置 </b></font>
      是指“已有网站页面中要放置广告的位置，用来放置通过本系统设置的广告位”。</p>
      <p>　</td>
    </tr>
  </table>
  </center>
</div>

    <!--#include file="boot.asp"-->