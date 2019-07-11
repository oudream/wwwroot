<%
'以下为可更改设置项
'======================================================================
'沸腾3AS流浪尘缘新闻系统数据表名
Const db_Attach_Table = "[attach]"		'粘贴[上传]文件信息表X
Const db_BigClass_Table = "[BigClass]"		'大类表X
Const db_Board_Table = "[board]"		'公告板信息表X
Const db_Dep_Table = "[dep]"			'部门信息表X
Const db_Link_Table = "[link]"			'友情链接信息表X
Const db_News_Table = "[News]"			'图文系统信息表X
Const db_NewsFile_Table = "[newsfile]"		'暂无调用-----------------
Const db_Review_Table = "[Review]"		'评论信息表X
Const db_SmallClass_Table = "[smallclass]"	'小类表X
Const db_Special_Table = "[Special]"		'专题信息表X
Const db_System_Table = "[System]"		'系统设置信息表X
Const db_Type_Table = "[type]"			'文件类型信息表X
Const db_UploadPic_Table = "[uploadpic]"	'上载文件过渡信息表
Const db_Vote_Table = "[vote]"			'投票信息表
Const db_Manage_Table = "[Admin]"		'后台配置管理用户表
'======================================================================
'沸腾3AS流浪尘缘新闻系统与被整合系统共用的用户数据表
Const db_User_Table = "[Admin]"		'与整合系统共用的用户表
						'动网6.1 为 [User]
						'动网7.0 为 [Dv_User]
						'不整合则值为 [ADMIN]或者[FT_User]
'======================================================================
'沸腾3AS流浪尘缘新闻系统与动网论坛共用的用户字段名
Const db_User_ID = "ID"				'ID		用户ID号
Const db_User_Name = "UserName"			'UserName	用户登录名
Const db_User_Sex = "sex"			'sex		用户性别[先生=1 女士=0]
Const db_User_Email = "email"			'email		用户电子邮箱
Const db_User_Password = "Passwd"		'PassWD	用户登录密码
Const db_User_Question = "question"		'question	用户问题
Const db_User_Answer = "answer"			'answer	用户回答
Const db_User_Face = "photo"			'photo		用户照片文件路径
Const db_User_RegDate = "dateandtime"		'dateandtime	用户注册时间
Const db_User_LoginTimes = "logins"		'logins	用户登录次数
Const db_User_LastLoginTime = "lastlogin"	'lastlogin	用户最近一次登录时间
Const db_User_LastLoginIP = "IP"		'IP		用户最近登录IP
Const db_User_birthday = "birthday"		'birthday	用户生日，包含年月日
Const db_User_FaceWidth = "UserWidth"		'用户头象宽度
Const db_User_FaceHeight = "UserHeight"		'用户头象高度

'======================================================================
'沸腾展望新闻系统管理用户字段名（一般不要修改）
Const db_ManageUser_ID = "ID"				'ID		管理用户ID号
Const db_ManageUser_Name = "UserName"			'UserName	管理用户登录名
Const db_ManageUser_Sex = "sex"			'sex		管理用户性别[先生=1 女士=0]
Const db_ManageUser_Email = "email"			'email		管理用户电子邮箱
Const db_ManageUser_Password = "Passwd"		'PassWD	管理用户登录密码
Const db_ManageUser_Question = "question"		'question	管理用户问题
Const db_ManageUser_Answer = "answer"			'answer	管理用户回答
Const db_ManageUser_Face = "photo"			'photo		管理用户照片文件路径
Const db_ManageUser_RegDate = "dateandtime"		'dateandtime	管理用户注册时间
Const db_ManageUser_LoginTimes = "logins"		'logins	管理用户登录次数
Const db_ManageUser_LastLoginTime = "lastlogin"	'lastlogin	管理用户最近一次登录时间
Const db_ManageUser_LastLoginIP = "IP"		'IP		管理用户最近登录IP
Const db_ManageUser_birthday = "birthday"		'birthday	管理用户生日，包含年月日
'----------------------------------------------------------------------
'沸腾3AS流浪尘缘新闻系统独有字段
Const db_User_birthyear = "birthyear"		'原生日年
Const db_User_birthmonth = "birthmonth"		'原生日月
Const ReadNews_CopyRight_Logo = "images/ft_logo.gif"	'文章阅读版权标识
'======================================================================
'综合设置项

Const Forcast_SN="Forcast2004-0001"	'安装序列号

dim UserTableType,FileUploadPath,BbsPath
UserTableType = "FeiTium"	   		'整合动网论坛则值为"Dvbbs"
						'不整合则值为"FeiTium"
						'---------------------------------
FileUploadPath = "uploadfile/"			'图文系统文件上传路径[拟在增加]
						'---------------------------------
BbsPath = "bbs/"				'BBS相对于本系统的目录,最后需加"/"
ScrollClick = "double"				'滚动栏自动滚动鼠标事件方式,"double"为双击自动滚动方式,
						'其它值则为左键单击拖动方式。
mouse_wheel_zoom="on"				'新闻图片鼠标滚轮缩放,"on"为启用缩放效果,"off"为不启用
TransShow="off"					'页面切换效果，"on"为启用切换效果,"off"为不启用
eWebEditor="0"					'是否启用eWebEditor来代替沸腾原编辑器，值为"1"时启用，其它值为不启用。
IsDebug="0"							'是否为调试模式，值为"1"时启用,其它值为不启用.
'=======================================================================
'以上为可更改设置项

'=======================================================================
'以下信息一般情况下无需更改
'=======================================================================
dim db_Sex_select , db_Birthday_Select , FaceUploadPath , db_Set_Table , db_Forum_UserNum , db_Form_lastUser

if UserTableType = "FeiTium" then
	'------------兼容选项------------		========================================
	db_Sex_Select = "FeiTium"			'性别字段类型选项，主要为兼容原图文系统与整合系统
							'不整合则值为 "FeiTium",为文本型
							'沸腾3AS流浪尘缘新闻系统原用户性别字段类型为文本型
							'---------------------------------------
	db_Birthday_Select = "FeiTium"			'生日字段兼容选项
							'不整合则值为 "FeiTium",为一个仅包含日期的文本型字段,
							'需与birthyear，birthmonth字段配合，形成完整的年月日
							'=======================================
	FaceUploadPath = "uploadfile/face/"		'头象上传目录,最后需加"/"
							'不整合则值为 "uploadfile/face/"
							'沸腾3AS流浪尘缘自身的头象路径
							'---------------------------------
else
	if UserTableType = "Dvbbs" then
		'沸腾3AS流浪尘缘新闻系统 V0.45Finish版与动网论坛共用数据库的BBS设置参数表
		db_Set_Table = "[Dv_Setup]"
							'------------DVBBS设置参数变量-------------
		db_Forum_UserNum = "Forum_UserNum"	'论坛总用户数
		db_Forum_lastUser = "Forum_lastUser"	'论坛设置表最新注册用户

		'------------兼容选项------------	========================================
		db_Sex_Select = "Number"		'性别字段类型选项，主要为兼容原图文系统与整合系统
							'动网系统该值为 "Number",为数字型
							'动网用户性别字段类型为数值型
							'---------------------------------------
		db_Birthday_Select = "Text"	'生日字段兼容选项
							'动网系统该值为 "Text",为包含年月日的一个文本型字段
							'=======================================
		FaceUploadPath = "UploadFace/"		'头象上传目录,最后需加"/"
							'整合动网则为"UploadFace/",相对于BBS安装目录
							'---------------------------------
		FaceDefault = "Images/userface/image1.gif"
							'用户注册（沸腾图文侧）默认头象文件，相对于 BbsPath 路径。
	end if
end if

''============================================
''以下为原沸腾3AS流浪尘缘新闻系统 V0.45版内容
''============================================
Set rs9 = Server.CreateObject("ADODB.Recordset")
sql9 ="SELECT * From "& db_System_Table &""
RS9.open sql9,Conn,3,3
xpurl = rs9("xpurl")
email = rs9("email")
copyright = rs9("copyright")
version = rs9("version")
ver = rs9("ver")
logo = rs9("logo")
logourl = rs9("logourl")
banner = rs9("banner")
bannerurl = rs9("bannerurl")
bgcolor = rs9("bgcolor")
bgimg = rs9("bgimg")
zs = rs9("zs")
zs1 = rs9("zs1")
zs2 = rs9("zs2")
reg = rs9("reg")
inputmodpwd = rs9("inputmodpwd")
upload = rs9("upload")
fabiao = rs9("fabiao")
fabiaomod = rs9("fabiaomod")
fabiaocheck = rs9("fabiaocheck")
gd1 = rs9("gd1")
gd2 = rs9("gd2")
checkmod = rs9("checkmod")
checkdel = rs9("checkdel")
smallgl = rs9("smallgl")
specialgl = rs9("specialgl")
system = rs9("system")
css = rs9("css")
init = rs9("init")
delreview = rs9("delreview")
shsmallgl = rs9("shsmallgl")
shspecialgl = rs9("shspecialgl")
shdelreview = rs9("shdelreview")
reviewable = rs9("reviewable")
modnewsable = rs9("modnewsable")
checkshow = rs9("checkshow")
showuser = rs9("showuser")
search = rs9("search")
showsearch = rs9("showsearch")
showspecial = rs9("showspecial")
showdata = rs9("showdata")
showauthor = rs9("showauthor")
showclub = rs9("showclub")
showlink = rs9("showlink")
showlinkmap = rs9("showlinkmap")
showvote = rs9("showvote")
showcount = rs9("showcount")
linkreg = rs9("linkreg")
linkpass = rs9("linkpass")
linkmana = rs9("linkmana")
votemana = rs9("votemana")
powermana = rs9("powermana")
showip = rs9("showip")
reviewcheck = rs9("reviewcheck")
reviewcheckshow = rs9("reviewcheckshow")
showyear = rs9("showyear")
showtime = rs9("showtime")
showclick = rs9("showclick")
shownum = rs9("shownum")
freetime = rs9("freetime")
uselevel = rs9("uselevel")
uploadtype = rs9("uploadtype")

top_news = rs9("top_news")
bigclassshownum = rs9("bigclassshownum")
top_sp = rs9("top_sp")
top_txt = rs9("top_txt")
top_img = rs9("top_img")
topuser = rs9("topuser")
linkshownum = rs9("linkshownum")
reviewnum = rs9("reviewnum")

topbg = rs9("topbg")
topfont = rs9("topfont")
t_bg = rs9("t_bg")
b_bg = rs9("b_bg")
T_m_bg = rs9("T_m_bg")
l_BG = rs9("l_BG")
l_top = rs9("l_top")
l_main = rs9("l_main")
m_BG = rs9("m_BG")
m_main = rs9("m_main")
m_top = rs9("m_top")
r_top = rs9("r_top")
r_BG = rs9("R_BG")
r_main = rs9("r_main")
border = rs9("border")

gg = rs9("gg")
gg1 = rs9("gg1")
jjgn = rs9("name")
PicNum = rs9("picnum")			'定义图片JS新闻调用数
newsNum = rs9("newsnum")		'定义一般JS新闻调用数
rs9.close
set rs9 = nothing
%>
