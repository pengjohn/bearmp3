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
<table border="0" width="760" cellspacing="1" cellpadding="1">
	<tr>
		<td width="70%" valign="top" align="center">
			<table width="100%">
				<tr>
					<td><br>
    				<%
    				select case request("list_mode")
    					case "ALL_ALBUM"
							sql="select * from album where (albumid mod 1000)<>0 order by singerid,albumname"
							titlename = "所有专辑列表"
							listalbum
    					case "NEW_ALBUM"
							sql="select top 100 * from album where (albumid mod 100<>0) order by album_id desc"
							titlename = "新增100张专辑"
							listalbum
						case "ALL_SONG"
							list_all_song
						case "ALL_SONG_LISTPRO"
							list_all_song_4listpro
						case "MISS_SONG"
							sql="select * from song,album where song.albumid=album.albumid and song.songfile='' order by album.albumid desc"
							titlename = "所缺歌曲"
							showsongs
						case "MISS_COVER"
							sql="select * from album where albumcover='images/nocover.gif' order by albumname" 
							titlename = "缺少CD专辑封面"
							listalbum
						case else

					end select
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


<%'--------------------------全部专辑列表-------------------------------------%>
<%
sub listalbum 
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then 
		response.write "<p align='center'>暂时没有收集</p>" 
	else 
%>
		<table border="0" width="100%" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" bordercolordark="#CCCCCC" bordercolorlight="#000000">
		<tr>
			<td><%=titlename%><br><br></td>
		</tr>
			<tr>
				<td colspan="2">
					<table border="1" width="100%" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#CCCCCC" bordercolorlight="#000000">
						<tr align="center">
							<td width="12%" height="22" background="images/dh-di.jpg">专辑编号</td>
							<td width="40%" background="images/dh-di.jpg">CD专辑</td>
						</tr>
						<%do while not rs.eof%>
						<tr align="center">
							<td><%=rs("albumid")%></a></td>
							<td align="left"><a href="default.asp?albumid=<%=rs("albumid")%>"><%=rs("albumname")%></a></td>
						</tr>
						<%
						rs.movenext
						loop
						%>
					</table>
				</td>
			</tr>
		</table>
<%	end if
	rs.close
	set rs=nothing
end sub
%>

<%'--------------------------新增专辑列表-------------------------------------%>
<%sub newalbum_show%>
    <table border="0" width="100%" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" bordercolordark="#CCCCCC" bordercolorlight="#000000">
		<tr>
			<td>当前位置：新增200张专辑列表<br><br></td>
		</tr>
		<tr>
			<td bgcolor="#E2F1F1">
				<font color="black">
		<%
		i=0
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1		
			do while not rs.eof%>
				专辑：<%=rs("albumname")%><br>
		<%
			i=i+1
			if i>=200 then exit do
		rs.movenext
		loop
		rs.close
		set rs=nothing
		%>
				</font>
			</td>
		</tr>
	</table>
<% end sub %>

<%'--------------------------歌曲列表-------------------------------------%>
<%
sub showsongs
	dim i 
	i=0 
	%>
	<table border="0" width="100%" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" bordercolordark="#CCCCCC" bordercolorlight="#000000">
		<tr>
			<td width="100%">当前位置：<%=titlename%>500首	</td>
		</tr>
		<tr>
			<td>
				<table border="1" width="100%" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#CCCCCC" bordercolorlight="#000000">
					<tr>
						<td width="35%" height="22" align="center" background="images/dh-di.jpg">歌曲名称</td>
						<td width="35%" height="22" align="center" background="images/dh-di.jpg">所属专辑</td>
					</tr>
					<%
					set rs=server.createobject("adodb.recordset")
					rs.open sql,conn,1,1
						do while not rs.eof
					%>
					<tr>
						<td>
							<%if rs("songhot")>=5 then%>
								<font color="red">★</font>
							<%end if%>
							<%if rs("songfile")<>"" then%>
								<a href="<%=rs("songfile")%>"><%=(i+1)%>. <%=rs("songname")%></a>
							<%else%>
								<%=(i+1)%>. <%=rs("songname")%>(缺)
							<%end if%>
						</td>
						<%if rs("albumid")=0 then%>
							<td align="center">会员上传</td>
						<%else%>
							<td align="center"><a href="default.asp?albumid=<%=rs("albumid")%>"><%=rs("albumname")%></a></td>
						<%end if%>
					</tr>
					<%
					i=i+1
					if i>=500 then exit do
					rs.movenext
					loop
					rs.close
					set rs=nothing
					%>
				</table>
			</td>
		</tr>
	</table>
<% end sub %>

<%'--------------------------歌曲列表-------------------------------------%>
<%
sub list_all_song
  Set fs=Server.CreateObject("Scripting.FileSystemObject")
  set ts = fs.createtextfile(Server.MapPath("\")&"\allsongs.txt",true) '写文件

	set rs=server.createobject("adodb.recordset")
	sql="select * from song,album where song.albumid=album.albumid and song.songfile<>'' order by songid desc"
	rs.open sql,conn,1,1

	albumid = 0
	do while not rs.eof
		if rs("albumid")<>albumid then
			ts.write "-----------------------------------------"&chr(13)
			ts.write rs("albumname")&chr(13)
			ts.write "-----------------------------------------"&chr(13)
			albumid = rs("albumid")
		end if
		ts.write (rs("songname"))&chr(13)
	rs.movenext
	loop
	rs.close
	set rs=nothing
  ts.close
  set ts=nothing    '文件生成
%>
<a href="allsongs.txt">下载歌曲列表</a>
<% end sub %>

<%'--------------------------歌曲列表-------------------------------------%>
<%
sub list_all_song_4listpro
'  Set fs=Server.CreateObject("Scripting.FileSystemObject")
'  set ts = fs.createtextfile(Server.MapPath("\")&"\allsongs_listpro.txt",true) '写文件

	set rs_singer=server.createobject("adodb.recordset")
	rs_singer.open "select * from singer where singer.classid = 1 order by singer.singername ", conn,1,1
	do while not rs_singer.eof
	singername = rs_singer("singername")
	singerid = rs_singer("singerid")
		response.write "0CHANGETAB"&rs_singer("singername")&"<br>"
		set rs=server.createobject("adodb.recordset")
		sql="select * from song,album where song.albumid=album.albumid and song.songfile<>'' and album.singerid="&singerid&" order by album.albumname desc,song.songid"
		rs.open sql,conn,1,1
		albumid = 0
		do while not rs.eof
			if rs("albumid")<>albumid then
				response.write """<br>"
				response.write "1CHANGETAB"&rs("albumname")&"CHANGETAB""<br>"
				albumid = rs("albumid")
			end if
			response.write (rs("songname"))&"<br>"
		rs.movenext
		loop
		rs.close
		set rs=nothing
	rs_singer.movenext
	loop
	rs_singer.close
	set rs_singer = nothing
	
'  ts.close
'  set ts=nothing    '文件生成
%>
<a href="allsongs_listpro.txt">下载歌曲列表</a>
<% end sub %>