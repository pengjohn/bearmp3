<!--#include file="conn_music_open.asp"-->
<!--#include file="conn_function.asp"-->
<%
dim song_id
dim songlrc
dim songfile
dim songname
dim lrcfile
dim initfloder

dim sql 
dim rs

song_id=request("song_id")

set rs=server.createobject("adodb.recordset")
sql="select * from song where song_id="&cstr(song_id)  
rs.open sql,conn,1,3
	songlrc=rs("songlrc")
	songfile=rs("songfile")
	songname=rs("songname")
rs.close
set rs=nothing

if songfile<>"" then
	lrcfile = replace(songfile,".mp3",".lrc")
	lrcfile = replace(lrcfile,".MP3",".lrc")
	lrcfile = replace(lrcfile,".Mp3",".lrc")
	lrcfile = replace(lrcfile,".mP3",".lrc")
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>骚熊音乐->歌曲歌词</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body bgcolor="#000000" text="#FFFFFF" link="#0000FF" vlink="#0000FF" alink="#FF0000">
<div align="center">
<table border="1" width="500" bordercolor="#FFFFFF">
	<tr>
		<td  align="center"><p>歌名：<%=songname%></p>
			<%if songlrc<>"" then
				Response.Write changechr(songlrc)
  		  	else
	  			Response.Write("歌词：无")
			end if%>
			<%
			dim fs
			Set fs = CreateObject("Scripting.FileSystemObject")
			if fs.FileExists(Server.MapPath("./")&"\"&lrcfile) then
				Response.Write("<br><br><a href='"&lrcfile&"'>下载lrc歌词</a>")
			end if
			%>
    	</td>
	</tr>
	<%if songlrc<>"" then%>
		<%if Session.Contents("IsAdmin")=true then%>
			<form method="POST" action="lrc_save.asp">
			<input type="hidden" name="song_id" size="15" value="<%=song_id%>">
			<input type="hidden" name="action" size="15" value="edit">
			<tr>
				<td width="80%" height="30"> <textarea rows="18" name="songlrc" cols="60"><%=songlrc%></textarea></td>
			</tr>
			<tr>
				<td valign="middle" align="center"> 
					<input type="submit" value=" 修 改 " name="cmdok"style="background-color: rgb(0,0,0); color: rgb(255,255,255); border: 1px dotted rgb(255,255,255)">&nbsp;
					<input type="reset"  value=" 恢 复 " name="cmdcancel" style="background-color: rgb(0,0,0); color: rgb(255,255,255); border: 1px dotted rgb(255,255,255)">
				</td>
			</tr>
			</form>
		<%end if%> 
	<%else%>
		<%if Session.Contents("mp3_username")<>"" then%>
			<tr>  
				<td>
					<form method="POST" action="lrc_save.asp">
					<input type="hidden" name="song_id" size="15" value="<%=song_id%>">
					<input type="hidden" name="action" size="15" value="addnew">
					<div align="center"><center>
					<table border="0" cellspacing="1" width="100%">
						<tr>
							<td width="20%" align="right" height="30">添加歌词：</td>
							<td width="80%" height="30"> <textarea rows="12" name="songlrc" cols="40"></textarea></td>
						</tr>
					</table>
					<input type="submit" value=" 添 加 " name="cmdok"style="background-color: rgb(0,0,0); color: rgb(255,255,255); border: 1px dotted rgb(255,255,255)">&nbsp;
					<input type="reset" value=" 清 除 " name="cmdcancel" style="background-color: rgb(0,0,0); color: rgb(255,255,255); border: 1px dotted rgb(255,255,255)">
					</form>
				</td>
			</tr>
		<%end if%>
	<%end if%>
</table>
</div>
</body>
</html>
<!--#include file="conn_music_close.asp"-->
