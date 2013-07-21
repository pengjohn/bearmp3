<%
	const	Info_Title				= "夏新音乐网"
	const Info_Release			="　　2005年新的开始。"		'首页通告
   const Info_Homepage			= "http://10.100.51.45/music"
   const Info_Email				= "mailto:pengjohn@amoi.com.cn"
   const Info_Version			= "Ver3.1"
   const Switch_Comment		= 0	'歌曲、专辑评论开关
   const Switch_DianGe		= 0	'点歌台开关
   const Switch_TopSong		= 1	'歌曲排行
   const Switch_TopAlbum	= 1	'专辑排行
   const Switch_GuestBook	= 1	'专辑排行
   
	const Mode_Play			= 1	'播放模式， 1-无歌词; 2-带歌词weblrc
	const Mode_Download		= 0	'下载模式， 0-直接下载; 1-转向下载

   const MaxPerPage_Songs		= 20	'显示总歌曲时，每页显示歌曲数
   const MaxPerPage_GstBook	= 30	'每页留言条数
   const MaxPerPage_User		= 30	'显示会员时，每页显示会员数
   const MaxPerPage_NewAlbum	= 5	'显示新增专辑时，每页显示专辑数
   const MaxPerRow_Singer			= 4	'显示歌手时，每行显示歌手数

   const Num_Gstbook			= 10	'首页最新留言条数
   const CommentNum   		= 10	'首页最新评论条数
   const Num_Goodcomment		= 1	'首页优秀评论条数
   const Num_TopSong				= 10	'歌曲排行版数

   const date_start 		=  #8/26/2001#		'开站时间
   const date_start_show 	= "2001年08月27日"	'显示开站时间

   const Mul_Class = 2		'类型2位
   const Mul_Singer= 3		'歌手3位
   const Mul_Album = 3		'专辑3位
   const Mul_Song  = 3		'歌曲3位
%>