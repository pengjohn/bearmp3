<!--#include file="conn_music_open.asp" -->  
<!--#include file="conn_function.asp" -->  
<%
dim rs,sql
dim commentname
dim commentemail
dim oommenttitle
dim commentcontent
dim commentindex

commentname=Session.Contents("mp3_username")
commentemail=""
commentcontent=request("commentcontent")
commentindex=request("commentindex")

set rs=server.createobject("adodb.recordset")
if (commentindex \ (10^(Mul_Class+Mul_Singer)))=0 then		'commenttitle="歌手"
    sql="select * from singer where singerid="+cstr(commentindex)
    rs.Open sql,conn,1,3
    commenttitle = "［歌手］"&rs("singername")
else
     	if (commentindex \ (10^(Mul_Class+Mul_Singer+Mul_Album)))=0 then	'commenttitle="专辑"
		sql="select * from album where albumid="+cstr(commentindex)
		rs.Open sql,conn,1,3
		commenttitle = "［专辑］"&rs("albumname")
     	else							'commenttitle="歌曲"
		sql="select * from song where songid="+cstr(commentindex)
		rs.Open sql,conn,1,3
		commenttitle = "［歌曲］"&rs("songname")
     	end if
end if
rs.close
set rs=nothing


if commentname="" then
	commentname="匿名"
end if
if commentemail="" then
	commentemail="Null"
end if
if commentcontent="" then
	commentcontent="无"
end if


set rs=server.createobject("adodb.recordset")
sql="select * from comment where (comment_id is null)"     
rs.Open sql,conn,1,3
rs.AddNew
	rs("commentdatetime")=now()
	rs("commentname")=commentname
	rs("commentemail")=commentemail
	rs("commentip")=Request.ServerVariables("REMOTE_ADDR")
	rs("commenttitle")=commenttitle
	rs("commentcontent")=commentcontent
	rs("commentindex")=commentindex
rs.Update
rs.close
set rs=nothing
score_add_comment	'加积分
%>
<!--#include file="conn_music_close.asp" -->  
<%Response.Redirect ("comment_show.asp?commentindex="&commentindex)%>