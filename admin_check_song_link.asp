<!--#include file="conn_music_open.asp"-->
<!--#include file="conn_define.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>骚熊音乐</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#000000">
<div align="center">
<table border="0" width="755" cellspacing="1" cellpadding="1">
	<tr>
		<td width="70%" valign="top" align="center">
			<table width="100%">
				<tr>
					<td><br>
    		<% '歌曲类型classid从00-21, 80-99为功能扩展 80-最新上传 81-歌曲搜索(歌曲) 84-歌曲搜索(歌词) 86-下载排行 87-好歌推荐 88-所缺歌曲 89-缺封面专辑
					listsongs
		%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</div>
<p></p>
<!--#include file="conn_music_close.asp"-->
</body>
</html>

<%
sub listsongs			
	set rs=server.createobject("adodb.recordset")
	sql="select * from song,album where song.albumid=album.albumid and album.classid=1 and song.songfile<>'' order by song_id desc"
	rs.open sql,conn,1,1

	dim fs
	i=1
	Set fs = CreateObject("Scripting.FileSystemObject")
	do while not rs.eof
		if fs.FileExists(Server.MapPath("./")&"\"&rs("songfile")) then
		else
			Response.Write(rs("songfile"))
			Response.Write("<br>")			
			i=i+1
		end if
	if i>=100 then exit do
	rs.movenext
	loop
	rs.close
end sub %>