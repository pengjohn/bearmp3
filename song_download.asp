<!--#include file="conn_music_open.asp"-->
<%
dim song_id
dim songfile
song_id=request("song_id")

set rs=server.createobject("adodb.recordset")
rs.open ("SELECT * FROM song WHERE song_id=" & song_id),conn,1,3
	songfile=rs("songfile")
	rs("songhit")=rs("songhit")+1
rs.Update
rs.close
set rs=nothing
%>
<!--#include file="conn_music_open.asp"-->
<%Response.Redirect (songfile)%>

