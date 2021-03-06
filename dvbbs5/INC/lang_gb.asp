<%
'=========================================================
' Language For Chinese GB Code
' File: lang_gb.asp
' Version:5.0
' Date: 2002-9-17
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

'First Msg
Const FirstLoginMsg="正在登陆论坛……<br><br>本系统要求使用COOKIES，假如您的浏览器禁用COOKIES，您将不能登录本系统……"
Const FirstIncMsgErr="错误的系统参数，很可能您指定了错误的论坛名称，请从有效连接进入。"
Const FirstIncReflashMsg="正在打开页面，请稍后……"
Const FirstIncIpErr="您的IP已经被限制不能访问本论坛，请和管理员联系。"
Const FirstIncStats="论坛信息"
Const Forum_charset="GB2312"

'顶部导航，部分通用信息
Function Language_var(selpage)
Dim Lang_nav(30)
Lang_nav(0)="欢迎您，"
Lang_nav(1)="请先登陆"
Lang_nav(2)="注册"
Lang_nav(3)="重登陆"
Lang_nav(4)="用户控制面板"
Lang_nav(5)="排行"
Lang_nav(6)="搜索"
Lang_nav(7)="帮助"
Lang_nav(8)="退出"
Lang_nav(9)="管理"
Lang_nav(10)="我参与的主题"
Lang_nav(11)="已被回复的我的发言"
Lang_nav(12)="您的帖子有人回复了"
Lang_nav(13)="您有新短消息"
Lang_nav(14)="[返回]"
Lang_nav(15)="注册用户"
Lang_nav(16)="忘记密码"
Lang_nav(17)="用户名："
Lang_nav(18)="密码："
Lang_nav(19)="不保存"
Lang_nav(20)="保存一天"
Lang_nav(21)="保存一月"
Lang_nav(22)="保存一年"
Lang_nav(23)="登 陆"
Lang_nav(24)="版本："
Lang_nav(25)="页面执行时间："
Lang_nav(26)="毫秒"
Lang_nav(27)="错误的版面参数！请确认您是从有效的连接进入。"
Lang_nav(28)="本论坛为认证论坛，请<a href=login.asp>登陆</a>并确认您的用户名已经得到管理员的认证后进入。"
Lang_nav(29)="本论坛为认证论坛，请确认您的用户名已经得到管理员的认证后进入。"
Lang_nav(30)="您没有浏览本论坛的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
Lang_nav(31)="关闭详细列表"
Lang_nav(32)="显示详细列表"
Lang_nav(33)="请指定论坛版面。"
Lang_nav(34)="非法的版面参数。"
Lang_nav(35)="请指定相关贴子。"
Lang_nav(36)="非法的贴子参数。"
Lang_nav(37)="您没有浏览在本论坛查看其他人发布的帖子的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
Lang_nav(38)="您指定的贴子不存在"
Lang_nav(39)="发表一个新主题"
Lang_nav(40)="发表一个新投票"
Lang_nav(41)="刷新"
Lang_nav(42)="发布小字报"
Lang_nav(43)="该用户发言已被管理员屏蔽"
Lang_nav(44)="您指定的贴子不存在"
Lang_nav(45)="本主题贴数"
Lang_nav(46)="回复主题"
Lang_nav(47)="跳转论坛至..."
Lang_nav(48)="没有论坛"
Lang_nav(49)="HTML标签："
Lang_nav(50)="UBB标签："
Lang_nav(51)="贴图标签："
Lang_nav(52)="Flash标签："
Lang_nav(53)="表情字符转换："
Lang_nav(54)="上传图片："
Lang_nav(55)="最多"
Lang_nav(56)="可用"
Lang_nav(57)="不可用"
Lang_nav(58)="可以使用Ctrl+Enter直接提交"
Lang_nav(59)="邮件回复"
Lang_nav(60)="显示签名"
Lang_nav(61)="清空内容！"
Lang_nav(62)="预 览"
Lang_nav(63)="OK!发表我的回应帖子"
Lang_nav(64)="清空内容！"
Lang_nav(65)="清空内容！"
Lang_nav(66)="清空内容！"
Lang_nav(67)="清空内容！"

Dim Lang_code(20)
Lang_code(0)=","
Lang_code(1)="，"
Lang_code(2)="."
Lang_code(3)="。"
Lang_code(4)="'"
Lang_code(5)="‘"
Lang_code(6)="’"
Lang_code(7)="""
Lang_code(8)="“"
Lang_code(9)="”"
Lang_code(10)="("
Lang_code(11)="（"
Lang_code(12)=")"
Lang_code(13)="）"
Lang_code(14)="["
Lang_code(15)="【"
Lang_code(16)="]"
Lang_code(17)="】"
Lang_code(18)="╋"
Lang_code(19)="├"

Select Case selpage
Case "Forum_index"
'首页
Dim Lang_index(50)
Lang_index(0)="论坛首页"
Lang_index(1)="欢迎新会员"
Lang_index(2)="新进来宾"
Lang_index(3)="今日贴数："
Lang_index(4)="主题总数："
Lang_index(5)="帖子总数："
Lang_index(6)="注册会员"
Lang_index(7)="管理团队"
Lang_index(8)="查看新贴"
Lang_index(9)="用户列表"
Lang_index(10)="发贴排行"
Lang_index(11)="加入收藏"
Lang_index(12)="联系我们"
Lang_index(13)="新手上路"
Lang_index(14)="论坛消息广播:"
Lang_index(15)="当前没有公告"
Lang_index(16)="-=> 快速登录入口"
Lang_index(17)="状态"
Lang_index(18)="论坛名称"
Lang_index(19)="版主"
Lang_index(20)="主题"
Lang_index(21)="贴子"
Lang_index(22)="最后发表"
Lang_index(23)="暂无"
Lang_index(24)="有新帖子"
Lang_index(25)="无新帖子"
Lang_index(26)="作者："
Lang_index(27)="转到："
Lang_index(28)="主题："
Lang_index(29)="论坛回收站"
Lang_index(30)="论坛所有版面版主删除的帖子。"
Lang_index(31)="-=> 友情论坛"
Lang_index(32)="-=> 用户来访信息"
Lang_index(33)="您的真实ＩＰ 是："
Lang_index(34)="-=> 论坛在线统计"
Lang_index(35)="同时在线峰值"
Lang_index(36)="发生时刻"
Lang_index(37)=""
Lang_index(38)=""
Lang_index(39)="在线用户"
Lang_index(40)="目前论坛上总共有 <b>"&clng(allonline())&"</b> 人在线，其中注册会员 <b>"&onlinenum&"</b> 人，访客 <b>"&guestnum&"</b> 人。"
Lang_index(41)="在线名单图例"
Lang_index(42)="总坛主"
Lang_index(43)="论坛坛主"
Lang_index(44)="论坛贵宾"
Lang_index(45)="普通会员"
Lang_index(46)="客人或隐身会员"

case "Forum_list"
Dim Lang_list(50)
Lang_list(0)="无版主"
Lang_list(1)="当前没有公告"
Lang_list(2)="查看所有的主题"
Lang_list(3)="查看一天内的主题"
Lang_list(4)="查看两天内的主题"
Lang_list(5)="查看一星期内的主题"
Lang_list(6)="查看半个月内的主题"
Lang_list(7)="查看一个月内的主题"
Lang_list(8)="查看两个月内的主题"
Lang_list(9)="查看半年内的主题"
Lang_list(10)="今日贴子"
Lang_list(11)=""
Lang_list(12)=""
Lang_list(13)=""
Lang_list(14)=""
Lang_list(15)="广播："
Lang_list(16)="精华"
Lang_list(17)="在线"
Lang_list(18)="事件"
Lang_list(19)="帮助"
Lang_list(20)="管理"
Lang_list(21)="状态"
Lang_list(22)="主 题"
Lang_list(23)="点<img src="&Forum_info(7)&"plus.gif>即可展开贴子列表"
Lang_list(24)="作 者"
Lang_list(25)="回复/人气"
Lang_list(26)="最后更新 | 回复人"
Lang_list(27)="本论坛暂无内容，欢迎发贴：）"
Lang_list(28)="固顶主题"
Lang_list(29)="精华帖子"
Lang_list(30)="投票贴子"
Lang_list(31)="主题已锁定"
Lang_list(32)="论坛已锁定"
Lang_list(33)="热门主题"
Lang_list(34)="开放主题"
Lang_list(35)="展开贴子列表"
Lang_list(36)="作者："
Lang_list(37)="发表于"
Lang_list(38)="最后跟贴："
Lang_list(39)="票"
Lang_list(40)="正在读取关于本主题的跟贴，请稍侯……"
Lang_list(41)="移动帖子请选择"
Lang_list(42)=""
Lang_list(43)="全选/取消"
Lang_list(44)="批量删除"
Lang_list(45)="批量移动"
Lang_list(46)="批量精华"
Lang_list(47)="批量锁定"
Lang_list(48)="批量固顶"
Lang_list(49)="执行"
Lang_list(50)="您确定执行的操作吗?"
Lang_list(51)="页次："
Lang_list(52)="页"
Lang_list(53)="每页"
Lang_list(54)="主题数"
Lang_list(55)="分页："
Lang_list(56)="首页"
Lang_list(57)="上十页"
Lang_list(58)="下十页"
Lang_list(59)="尾页"
Lang_list(60)="转到:"
Lang_list(61)=""
Lang_list(62)="图例"
Lang_list(63)="所有时间均为"
Lang_list(64)="回复超过10贴"

case "Forum_dispbbs"
Dim Lang_dispbbs(50)
Lang_dispbbs(0)="浏览帖子"
Lang_dispbbs(1)="平板"
Lang_dispbbs(2)="树形"
Lang_dispbbs(3)="该帖子已经被管理员删除！"
Lang_dispbbs(4)="您还没有登陆，不能进行投票；或者已经过了投票期限。"
Lang_dispbbs(5)="查看投票用户"
Lang_dispbbs(6)="投 票"
Lang_dispbbs(7)="截止时间："
Lang_dispbbs(8)="您已经投过票了，请看结果吧。"
Lang_dispbbs(9)="您是本帖的第 <B>{viewnum}</B> 个阅读者"
Lang_dispbbs(10)="浏览上一篇主题"
Lang_dispbbs(11)="{skiname}方式浏览贴子"
Lang_dispbbs(12)="浏览下一篇主题"
Lang_dispbbs(13)="贴子主题:"
Lang_dispbbs(14)="保存该页为文件"
Lang_dispbbs(15)="报告本帖给版主"
Lang_dispbbs(16)="显示可打印的版本"
Lang_dispbbs(17)="把本贴打包邮递"
Lang_dispbbs(18)="把本贴加入论坛收藏夹"
Lang_dispbbs(19)="发送本页面给朋友"
Lang_dispbbs(20)="把本贴加入IE收藏夹"
Lang_dispbbs(21)="头衔："
Lang_dispbbs(22)="等级："
Lang_dispbbs(23)="威望："
Lang_dispbbs(24)="文章："
Lang_dispbbs(25)="积分："
Lang_dispbbs(26)="门派："
Lang_dispbbs(27)="注册："
Lang_dispbbs(28)="给{username}发送一个短消息"
Lang_dispbbs(29)="把{username}加入好友"
Lang_dispbbs(30)="查看{username}的个人资料"
Lang_dispbbs(31)="搜索{username}在{boardname}的所有贴子"
Lang_dispbbs(32)="点击这里发送电邮给{username}"
Lang_dispbbs(33)="访问{username}的主页"
Lang_dispbbs(34)="引用回复这个贴子"
Lang_dispbbs(35)="回复这个贴子"
Lang_dispbbs(36)="该用户发言已被管理员屏蔽"
Lang_dispbbs(37)="该用户不在线<br>发贴IP："
Lang_dispbbs(38)="编辑这个贴子"
Lang_dispbbs(39)="同意该帖观点，给他一朵鲜花，将消耗您10点金钱"
Lang_dispbbs(40)="鲜花"
Lang_dispbbs(41)="不同意该帖观点，给他一个鸡蛋，将消耗您10点金钱"
Lang_dispbbs(42)="鸡蛋"
Lang_dispbbs(43)="精华"
Lang_dispbbs(44)="解除精华"
Lang_dispbbs(45)="删除跟帖"
Lang_dispbbs(46)="复制"
Lang_dispbbs(47)="复制单个贴子到别的版面"
Lang_dispbbs(48)="本主题贴数"
Lang_dispbbs(49)="分页："
Lang_dispbbs(50)="*树形目录"
Lang_dispbbs(51)="顶端"
Lang_dispbbs(52)="回复："
Lang_dispbbs(53)="主题："
Lang_dispbbs(54)="*快速回复："
Lang_dispbbs(55)="内容"
Lang_dispbbs(56)="固顶"
Lang_dispbbs(57)="解固"
Lang_dispbbs(58)="锁定"
Lang_dispbbs(59)="解锁"
Lang_dispbbs(60)="删除主题"
Lang_dispbbs(61)="移动"
Lang_dispbbs(62)="奖励"
Lang_dispbbs(63)="惩罚"
Lang_dispbbs(64)="发布公告"
Lang_dispbbs(65)="对发起主题用户奖励"
Lang_dispbbs(66)="对发起主题用户惩罚"
Lang_dispbbs(67)=""
Lang_dispbbs(68)=""
Lang_dispbbs(69)=""
Lang_dispbbs(70)=""
End Select
End Function()
%>