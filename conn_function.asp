<%
   const ScoreAddSong 		= 10	'???��?��?��
   const ScoreAddLrc 		= 2	'???��?���䨺
   const ScoreAddCDCover 	= 10	'???��CD��a??
   const ScoreAddComment 	= 2	'???��?��?��
   const ScoreAddGoodComment= 40	'?-��䨮?D?��?��??��??
   const ScoreCopyGoodComment 	= 5	'��a??��?D?��?��??��??
   
   const ScoreDelSong 		= -5	'1����a��?��?��?��D��??��䨪?����??��?��
   const ScoreDelLrc 		= -2	'��?��?�䨪?����??���䨺
   const ScoreDelComment 	= -4	'��D????1��???��������??��??
   const ScoreDelGoodComment	= -20	'��D????3-???��������??-��䨰?��??��??


function changechr(str) 
    changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;") 
    changechr=replace(replace(replace(replace(changechr,"[img]","<img src="),"[b]","<b>"),"[red]","<font color=CC0000>"),"[big]","<font size=7>") 
    changechr=replace(replace(replace(replace(changechr,"[/img]","></img>"),"[/b]","</b>"),"[/red]","</font>"),"[/big]","</font>") 

    changechr=replace(changechr,"[url]", "<a href=")
    changechr=replace(changechr,"[title]", ">")
    changechr=replace(changechr,"[/url]", "</a>")

    changechr=replace(changechr,"[big]", "<big>")
    changechr=replace(changechr,"[/big]", "</big>")

    changechr=replace(changechr,"[flash]", "<OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=600 height=600><PARAM NAME=movie VALUE=")
    changechr=replace(changechr,"[/flash]", "><PARAM NAME=quality VALUE=high></OBJECT>")

    changechr=replace(changechr,"[music]", "<object id=MediaPlayer1 height=145 width=388 classid=clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95><param name=AudioStream value=-1><param name=AutoStart value=-1><param name=AllowScan value=-1><param name=BufferingTime value=5><param name=Filename value=")
    changechr=replace(changechr,"[/music]","><param name=ShowDisplay value=-1><param name=ShowGotoBar value=0><param name=ShowPositionControls value=-1><param name=ShowStatusBar value=-1><param name=TransparentAtStart value=0><param name=Volume value=0></object>")
end function 

function changechr_level1(str) 
    changechr_level1=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;") 
end function

function score_add(score)
   dim sql,rs,username
   username = Session.Contents("mp3_username")
   
   set rs=server.createobject("adodb.recordset")
   sql="select * from user where username='"&username&"'" 
   rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		errmsg=errmsg					'�޴�user
	else
		rs("userscore") = rs("userscore") + score	'��user����
		rs.update
	end if
   rs.close
   set rs=nothing
end function
'------------------------------------
function score_add_song
	score_add ScoreAddSong
end function
'------------------------------------
function score_add_lrc
	score_add ScoreAddLrc
end function
'------------------------------------
function score_add_cdcover
	score_add ScoreAddCDCover
end function
'------------------------------------
function score_add_comment
	score_add ScoreAddComment
end function
'------------------------------------
function score_add_goodcomment
	score_add ScoreAddGoodComment
end function
'------------------------------------
function score_copy_goodcomment
	score_add ScoreCopyGoodComment
end function
'------------------------------------
function score_del_song
	score_add ScoreDelSong
end function
'------------------------------------
function score_del_lrc
	score_add ScoreDelLre
end function
'------------------------------------
function score_del_comment
	score_add ScoreDelComment
end function
'------------------------------------
function score_del_goodcomment
	score_add ScoreDelGoodComment
end function
'------------------------------------
%>



