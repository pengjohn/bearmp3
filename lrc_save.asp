<!--#include file="conn_music_open.asp"-->
<!--#include file="conn_function.asp"-->
<%
dim song_id
dim songlrc
dim action

dim rs
dim sql

song_id = request("song_id")
songlrc = request("songlrc")
action = request("action")

set rs=server.createobject("adodb.recordset")
sql="select * from song where song_id="+cstr(song_id)+" " 
rs.open sql,conn,1,3

select case action
	case "edit"
		rs("songlrc") = songlrc&chr(13)&"◇歌词由"&Session.Contents("mp3_username")&"修改◇"
	case "addnew"
		rs("songlrc") = songlrc&chr(13)&chr(13)&"◇歌词由"&Session.Contents("mp3_username")&"提供◇"
end select

rs("songdatetime") = date()
rs.update
rs.close
set rs=nothing
score_add_lrc		'加积分
%>
<!--#include file="conn_music_close.asp"-->
<%Response.Redirect ("lrc_show.asp?song_id="&song_id)%>

