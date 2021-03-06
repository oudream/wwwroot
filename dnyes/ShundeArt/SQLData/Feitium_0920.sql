if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Admin]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Admin]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BigClass]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BigClass]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Counters]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Counters]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FT_User]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FT_User]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[News]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[News]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Review]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Review]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Special]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Special]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[System]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[System]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Type]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Type]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Uploadpic]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Uploadpic]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Vote]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Vote]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[attach]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[attach]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[board]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[board]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[dep]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[dep]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[link]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[link]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[newsfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[newsfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[smallclass]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[smallclass]
GO

CREATE TABLE [dbo].[Admin] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[PassWD] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[purview] [int] NULL ,
	[OSKEY] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[fullname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[question] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[answer] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[sex] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[birthyear] [int] NULL ,
	[birthmonth] [int] NULL ,
	[birthday] [int] NULL ,
	[email] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[content] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[IP] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[number] [int] NULL ,
	[logins] [int] NULL ,
	[lastlogin] [smalldatetime] NULL ,
	[dateandtime] [smalldatetime] NULL ,
	[depname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[depid] [int] NULL ,
	[deptype] [int] NULL ,
	[adder] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[tel] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[shenhe] [int] NULL ,
	[jingyong] [int] NULL ,
	[reglevel] [int] NULL ,
	[photo] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[BigClass] (
	[BigClassID] [int] IDENTITY (1, 1) NOT NULL ,
	[BigClassMaster] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[bigclasszs] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[BigClassView] [int] NULL ,
	[BigClassName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[typeid] [int] NULL ,
	[bigclassorder] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Counters] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[total] [int] NOT NULL ,
	[today] [int] NOT NULL ,
	[yesterday] [int] NOT NULL ,
	[month] [int] NOT NULL ,
	[bmonth] [int] NOT NULL ,
	[date] [smalldatetime] NOT NULL ,
	[lastip] [nchar] (15) COLLATE Chinese_PRC_CI_AS NULL ,
	[inputdate] [smalldatetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FT_User] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[PassWD] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[purview] [int] NULL ,
	[OSKEY] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[fullname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[question] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[answer] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[sex] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[birthyear] [int] NULL ,
	[birthmonth] [int] NULL ,
	[birthday] [int] NULL ,
	[email] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[content] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[IP] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[number] [int] NULL ,
	[logins] [int] NULL ,
	[lastlogin] [smalldatetime] NULL ,
	[dateandtime] [smalldatetime] NULL ,
	[depname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[depid] [int] NULL ,
	[deptype] [int] NULL ,
	[adder] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[tel] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[shenhe] [int] NULL ,
	[jingyong] [int] NULL ,
	[reglevel] [int] NULL ,
	[photo] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[News] (
	[NewsID] [int] IDENTITY (1, 1) NOT NULL ,
	[Title] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[checkked] [int] NULL ,
	[Author] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[editor] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Original] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[image] [int] NULL ,
	[UpdateTime] [smalldatetime] NULL ,
	[Content] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[goodnews] [int] NULL ,
	[about] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[specialid2] [int] NULL ,
	[SpecialID] [int] NULL ,
	[EnCode] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[click] [int] NULL ,
	[typeid] [int] NULL ,
	[istop] [int] NULL ,
	[bigclassid] [int] NULL ,
	[smallclassid] [int] NULL ,
	[picnews] [int] NULL ,
	[picname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[newslevel] [int] NULL ,
	[titletype] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[titleface] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[titlecolor] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[titlesize] [int] NULL ,
	[newsrelated] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Review] (
	[ReViewID] [int] IDENTITY (1, 1) NOT NULL ,
	[NewsID] [int] NULL ,
	[Content] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[editor] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Author] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[passed] [int] NULL ,
	[reviewip] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[UpdateTime] [smalldatetime] NULL ,
	[Email] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[from] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[face] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[homepage] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[shengfen] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[oicq] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[reply] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[title] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Special] (
	[SpecialID] [int] IDENTITY (1, 1) NOT NULL ,
	[Specialmaster] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[specialzs] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[SpecialName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[System] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[name] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[xpurl] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[email] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[copyright] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[version] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[ver] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[logo] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[logourl] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[banner] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[zs] [int] NULL ,
	[zs2] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[zs1] [int] NULL ,
	[smallgl] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[specialgl] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[usecheck] [int] NULL ,
	[upload] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[bannerurl] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[bgcolor] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[bgimg] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[picnum] [int] NULL ,
	[showvote] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[showclub] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[showcount] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[linkreg] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[showlink] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[linkshownum] [int] NULL ,
	[showlinkmap] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[showuser] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[showauthor] [int] NULL ,
	[linkpass] [int] NULL ,
	[shownum] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[showclick] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[showyear] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[showtime] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[showdata] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[showspecial] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[showsearch] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[search] [int] NULL ,
	[reviewnum] [int] NULL ,
	[topuser] [int] NULL ,
	[newsnum] [int] NULL ,
	[top_news] [int] NULL ,
	[top_sp] [int] NULL ,
	[top_txt] [int] NULL ,
	[top_img] [int] NULL ,
	[T_BG] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[B_BG] [int] NULL ,
	[L_BG] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[L_TOP] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[L_MAIN] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[M_BG] [int] NULL ,
	[M_TOP] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[M_MAIN] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[R_BG] [int] NULL ,
	[R_TOP] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[R_MAIN] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[T_M_BG] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[topfont] [int] NULL ,
	[topbg] [int] NULL ,
	[border] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[reviewcheckshow] [int] NULL ,
	[reviewcheck] [int] NULL ,
	[showip] [int] NULL ,
	[linkmana] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[votemana] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[powermana] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[system] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[css] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[init] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[delreview] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[shsmallgl] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[shspecialgl] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[shdelreview] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[modnewsable] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[gd1] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[gd2] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[reg] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[fabiaomod] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[fabiaocheck] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[fabiao] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[checkdel] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[checkshow] [int] NULL ,
	[checkmod] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[reviewable] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[gg1] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[gg] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[bigclassshownum] [int] NULL ,
	[inputmodpwd] [int] NULL ,
	[freetime] [int] NULL ,
	[uselevel] [int] NULL ,
	[uploadtype] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[smallmana] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Type] (
	[typeid] [int] IDENTITY (1, 1) NOT NULL ,
	[typename] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[typecontent] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[mode] [int] NULL ,
	[typeorder] [int] NULL ,
	[typemaster] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[typeview] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Uploadpic] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[picname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[username] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Vote] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Title] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[select1] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[answer1] [int] NULL ,
	[select2] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[answer2] [int] NULL ,
	[select3] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[answer3] [int] NULL ,
	[select4] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[answer4] [int] NULL ,
	[select5] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[answer5] [int] NULL ,
	[select6] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[answer6] [int] NULL ,
	[select7] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[answer7] [int] NULL ,
	[select8] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[answer8] [int] NULL ,
	[DateAndTime] [smalldatetime] NULL ,
	[IsChecked] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[attach] (
	[attachid] [int] IDENTITY (1, 1) NOT NULL ,
	[newsid] [int] NULL ,
	[attachname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[attachcontent] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[attachtype] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[filename] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[board] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[title] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[content] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[inuse] [int] NULL ,
	[dateandtime] [smalldatetime] NULL ,
	[upload] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[dep] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[depname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[deptype] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[link] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[webname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[weburl] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[logo] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[webmaster] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[email] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[content] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[pass] [int] NULL ,
	[linktype] [int] NULL ,
	[dateandtime] [smalldatetime] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[newsfile] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[newsrelated] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[uploadfile] [ntext] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[smallclass] (
	[SmallClassID] [int] IDENTITY (1, 1) NOT NULL ,
	[smallclasszs] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[BigClassID] [int] NULL ,
	[smallclassma] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[smallmaster] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[typeid] [int] NULL ,
	[smallclassorder] [int] NULL ,
	[smallclassname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[smallclassview] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

/****** 系统初始脚本 ******/
/****** 2004-7-20 16:00:00 ******/

insert into Admin (UserName,PassWD,purview,OSKEY,fullname,question,answer,sex,birthyear,birthmonth,birthday,email,content,IP,number,logins,lastlogin,dateandtime,depname,depid,deptype,adder,tel,shenhe,jingyong,reglevel,photo)  values('base','db694daaa7d778ba','99999','super','小费','base','db694daaa7d778ba','先生','1980','01','01','base0321@163.com','系系统初始管理员','10.10.10.10','0','0','2004-7-20 16:00:00','2004-7-20 16:00:00','沸腾工作室','1','1','base','12345678','1','1','','uploadfile/face/200311120582086527.jpg');

insert into FT_User (UserName,PassWD,purview,OSKEY,fullname,question,answer,sex,birthyear,birthmonth,birthday,email,content,IP,number,logins,lastlogin,dateandtime,depname,depid,deptype,adder,tel,shenhe,jingyong,reglevel,photo)  values('base','db694daaa7d778ba','99999','super','小费','base','db694daaa7d778ba','先生','1980','01','01','base0321@163.com','系系统初始管理员','10.10.10.10','0','0','2004-7-20 16:00:00','2004-7-20 16:00:00','沸腾工作室','1','1','base','12345678','1','1','','uploadfile/face/200311120582086527.jpg');

insert into dep (depname,deptype) values ('沸腾工作室','1');

insert into link (webname,weburl,logo,webmaster,email,content,pass,linktype,dateandtime) values ('沸腾工作室','http://feitium.yeah.net','http://','base','base0321@163.com','沸腾工作室','1','0','2004-7-20 16:00:00');

insert into [System] (name,xpurl,email,copyright,version,ver,logo,logourl,banner,zs,zs2,zs1,smallgl,specialgl,usecheck,upload,bannerurl,bgcolor,bgimg,picnum,showvote,showclub,showcount,linkreg,showlink,linkshownum,showlinkmap,showuser,showauthor,linkpass,shownum,showclick,showyear,showtime,showdata,showspecial,showsearch,search,reviewnum,topuser,newsnum,top_news,top_sp,top_txt,top_img,T_BG,B_BG,L_BG,L_TOP,L_MAIN,M_BG,M_TOP,M_MAIN,R_BG,R_TOP,R_MAIN,T_M_BG,topfont,topbg,border,reviewcheckshow,reviewcheck,showip,linkmana,votemana,powermana,system,css,init,delreview,shsmallgl,shspecialgl,shdelreview,modnewsable,gd1,gd2,reg,fabiaomod,fabiaocheck,fabiao,checkdel,checkshow,checkmod,reviewable,gg1,gg,bigclassshownum,inputmodpwd,freetime,uselevel,uploadtype,smallmana) values ('沸腾工作室·展望新闻多媒体系统 V1.1 SQL版','feitium.yeah.net','base0321@163.com','沸腾展望新闻多媒体系统(核心:尘缘雅境图文系统)','V1.1 MSSQL版','Build1','images/f2.gif','http://feitium.yeah.net','images/f3.gif','1','1','1','0','0','1','0','http://feitium.yeah.net','#FFFFFF','','5','1','1','1','1','1','5','1','1','1','1','','1','0','1','1','1','1','1','5','5','5','5','5','5','5','1','0','#efefef','#76B0CE','http://feitium.yeah.net','0','#cccccc','#FFFFFF','1','./IMAGES/float.gif','快乐英语网，后台可设置关闭','','1','0','#C0C0C0','1','1','1','0','0','0','0','0','0','0','0','0','1','1','1','1','1','1','0','1','1','1','1','1','测试：用户名：base密码：feitium','','5','0','5','1','','0');
