<!--#include file="conn_music_open.asp"-->
<%
dim song_id
dim rs,sql
dim songid,songname,songfile

song_id = request("song_id")

set rs=server.createobject("adodb.recordset")
sql="select * from song where song_id="+cstr(song_id)+"" 
rs.open sql,conn,1,3
	songid=rs("songid")
	songname=rs("songname")
	songfile=rs("songfile")
rs.close
set rs=nothing
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>骚熊音乐->编辑歌曲</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body bgcolor="#000000">
<form method="POST" action="song_edit_save.asp">
<input type="hidden" name="song_id" size="15" value="<%=song_id%>">
<div align="center"><center>
<table border="1" cellspacing="0" width="90%" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0">
	<tr>
		<td width="100%" height="20"><div align="center"><center><p><b>歌曲信息编辑</b></td>
	</tr>
	<tr align="center">
		<td width="100%">
			<table border="0" cellspacing="1" width="100%">
				<tr>
					<td width="10%" align="right" height="30">歌曲编号：</td>
					<td width="90%" height="30"><input type="text" name="songid" size="25" value='<%=songid%>'></td>
				</tr>
				<tr>
					<td width="10%" align="right" height="30" >歌曲名称：</td>
					<td width="90%" height="30"><input type="text" name="songname" size="25" value='<%=songname%>'></td>
				</tr>
				<tr>
					<td width="10%" align="right" height="30">歌曲文件：</td>
					<td width="90%" height="30"><input type="text" name="songfile" size="80" value='<%=songfile%>'></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</center></div>
<div align="center"><center><p>
	<input type="submit" value=" 修 改 " name="cmdok"style="background-color: rgb(0,0,0); color: rgb(255,255,255); border: 1px dotted rgb(255,255,255)">&nbsp;
	<p><a href="song_edit.asp?song_id=<%=song_id%>">恢复</a></p>
</center></div>
</form>
<br><br>
</body>
</html>
<!--#include file="conn_music_close.asp"-->
