<%' step: CLASS -> SINGER -> ALBUM -> SAVEALBUM(auto add song) -> SONG%>

<!--#include file="conn_music_open.asp"-->
<!--#include file="conn_define.asp"-->
<%set rs=server.createobject("adodb.recordset")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>骚熊音乐-&gt;添加专辑</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#000000">
<%
select case request("step")
	case "CLASS"
%>
		<form method="POST" action="album_add.asp?step=SINGER">
		<div align="center"><center>
		<table border="0" cellspacing="10" width="500" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0">
			<tr>
				<td width="100%" height="20">第一步：专辑添加 → 归类</td>
			</tr>
			<tr align="center">
				<td width="100%">
					<img src="images/button-ff4.gif" alt WIDTH="28" HEIGHT="28">  请选择下属类型：
					<select name="classid" size="1">
						<%
						sql="select * from class"
						rs.open sql,conn,1,1
						Do until rs.eof
           				response.write "<option value='"&rs("classid")&"'>= "&rs("classname")&" =</option>"
						rs.MoveNext
						Loop
						rs.close
						%>
					</select>
				</td>
			</tr>
		</table>
		</center></div>
		<div align="center"><center><p>
			<input type="Image" src="images/button-ff.gif" tppabs title="下一步" name="Submit" WIDTH="73" HEIGHT="19">
		</p></center></div>
		</form>
	<%
	case "SINGER"
		classid = request("classid")
	%>
		<form method="POST" action="album_add.asp?step=ALBUM">
		<div align="center"><center>
		<table border="0" cellspacing="10" width="500" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0">
			<tr>
				<td width="100%" height="20">第二步：专辑添加 → 歌手：</td>
			</tr>
			<tr align="center">
				<td width="100%"><img src="images/button-ff4.gif" alt WIDTH="28" HEIGHT="28">  请选择歌手：
					<select name="singerid" size="1">
					<%
					sql="select * from singer where classid="+cstr(classid)+" order by singername desc"
					rs.open sql,conn,1,1
					Do until rs.eof
						response.write "<option value='"&rs("singerid")&"'>= "&rs("singername")&" =</option>"
					rs.MoveNext
					Loop
					rs.close
					%>
					</select>
					<br><br>若歌手不在列表中，请<a href="singer_add.asp?classid=<%=classid%>">添加歌手</a>
				</td>
			</tr>
		</table>
		</center></div>
		<div align="center"><center><p>
			<a href="#" onClick="history.back()"><img src="IMAGES/button-fb.gif" alt="上一步" border="0" WIDTH="73" HEIGHT="19"></a>&nbsp;&nbsp;
			<input type="Image" src="images/button-ff.gif" tppabs title="下一步" name="Submit" WIDTH="73" HEIGHT="19">
		</p></center></div>
		</form>
	<%
	case "ALBUM"
		singerid = request("singerid")
	%>
		<form method="POST" action="album_add.asp?step=SAVEALBUM">
		<input type="hidden" name="singerid" size="15" value="<%=singerid%>">
		<div align="center"><center>
		<table border="0" cellspacing="10" width="500" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0">
			<tr>
				<td width="100%" height="20">第三步：专辑添加 → 专辑信息：</td>
			</tr>
			<tr align="center">
				<td width="100%">
					<table border="0" cellspacing="1" width="100%">
						<tr>
							<td width="20%" align="right" height="30">专辑名称：</td>
							<td width="80%" height="30">
								<input type="text" name="albumname" size="25" maxlength="255">
								&nbsp;&nbsp;<input type="reset" value=" 清除输入 " name="cmdcancel" style="background-color: rgb(255,255,255); border: 1px inset rgb(0,0,0)">
							</td>
						</tr>
						<input type="hidden" name="albumgrade" value="1">
<%
'						<tr>
'							<td width="20%" align="right" valign="top" height="20">专辑评分：</td>
'							<td width="80%">
'								<select name="albumgrade" size="1">
'									<option value="1">　1　</option>
'									<option value="2">　2　</option>
'									<option value="3">　3　</option>
'									<option value="4">　4　</option>
'									<option value="5">　5　</option>
'								</select>
'							</td>
'						</tr>
%>
						<tr>
							<td width="20%" align="right" height="30">歌名模式：</td>
							<td width="80%" height="30">
								<input type=radio value=1 name=file_mode checked>歌手01 - 歌名.mp3<br>
<%'								<input type=radio value=2 name=file_mode> 01_歌手 - 歌名.mp3%>
							</td>
						</tr>
						<tr>
							<td width="20%" align="right" height="30">专辑封面：</td>
							<td width="80%" height="30">文件名同专辑名，扩展名为"jpg"或"gif"</td>
						</tr>
						<tr>
							<td width="20%" align="right" valign="top">已有专辑：</td>
							<td width="80%">
							<%
							sql="select * from album where singerid="+cstr(singerid)+" and ((albumid mod 100)<>0) order by albumname"
							rs.open sql,conn,1,1
							Do until rs.eof
								response.write rs("albumname")&"<br>"
							rs.MoveNext
							Loop
							rs.close
							%>							
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		</center></div>
		<div align="center"><center>
			<a href="#" onClick="history.back()"><img src="IMAGES/button-fb.gif" alt="上一步" border="0" WIDTH="73" HEIGHT="19"></a>&nbsp;&nbsp;
			<input type="Image" src="images/button-ff.gif" tppabs title="下一步" name="Submit" WIDTH="73" HEIGHT="19">
		</center></div>
		</form>
	<%
	case "SAVEALBUM"
		singerid = request("singerid")
		albumname = request("albumname")
		albuminfo = request("albuminfo")
		albumgrade = request("albumgrade")
		classid = singerid \ 10^Mul_Singer

		founerr=false
		if trim(request.form("albumname"))="" then
			founderr=true
			errmsg=errmsg+"<li>专辑名字不能为空</li>"
		end if
		'判断是否有同名专辑, 若专辑名中带"'",则跳过
		if instr(albumname,"'")=0 then	 '判断专辑名中是否有带"'"
			sql="select * from album where albumname='"&albumname&"'" 
			rs.open sql,conn,1,3
				if rs.eof and rs.bof then
					errmsg=errmsg
				else
					founderr=true
					errmsg=errmsg+"<li>已有同名专辑</li>"
				end if
			rs.close
		end if
		'取得类型名字
		sql="select * from class where classid="+cstr(classid)+"" 
		rs.open sql,conn,1,3
			classname=rs("classname")
		rs.close
		'取得歌手名字
		sql="select * from singer where singerid="+cstr(singerid)+"" 
		rs.open sql,conn,1,3
			singername=rs("singername")
		rs.close

		if founderr=false then
		'如果歌手有专辑则专辑id+1, 若歌手无专辑则专辑id=1
			sql="select * from album where singerid="+cstr(singerid)+" order by albumid desc" 
			rs.open sql,conn,1,3
				if rs.eof and rs.bof then
					albumid=singerid*(10^Mul_Album)+1
				else
					albumid=rs("albumid")+1
				end if
			rs.close
			
			'自动判断封面文件 *.jpg -> *.gif -> nocover.gif
			Set fs = CreateObject("Scripting.FileSystemObject")
			albumcover = "songs/"&classname&"/"&singername&"/"&albumname&"/"&albumname&".jpg"
			if fs.FileExists(Server.MapPath("./")&"\"&albumcover) then
			else
				albumcover = "songs/"&classname&"/"&singername&"/"&albumname&"/"&albumname&".gif"
				if fs.FileExists(Server.MapPath("./")&"\"&albumcover) then
				else
					albumcover="images/nocover.gif"
				end if
			end if

			'添加专辑资料到数据库
			sql="select * from album where (albumid is null)" 
			rs.open sql,conn,1,3
			rs.addnew
				rs("albumid")  = albumid
				rs("albumname")= albumname
				rs("albumcover") = albumcover
				rs("singerid") = singerid
				rs("classid") = classid
				rs("albumdatetime") = date()
				rs("albuminfo") = albuminfo
				rs("albumgrade") = albumgrade	
			rs.update
			rs.close
	%>
			<div align="center"><center>
			<table border="0" cellspacing="10" width="500" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0">
				<tr>
					<td>第四步：专辑添加 → 添加专辑成功：</td>
				</tr>
				<tr>
					<td width="100%">
						新增专辑：<%=albumname%><br>
						⒈ 新增目录：<br>   
							<%
							dim fs
							Set fs = CreateObject("Scripting.FileSystemObject")
							tmpfolder=Server.MapPath("./") & "\songs\"&classname&"\"&singername&"\"&albumname&"\"
						
							if not fs.FolderExists(tmpfolder) then
								fs.CreateFolder tmpFolder
								Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;"&tmpfolder&"<br>")
							end if
							Response.Write "⒉ 专辑封面：<br><img src='"& albumcover &"'><br>"
							Response.Write "⒉ 自动添加以下歌曲：<br>"

'自动添加专辑目录下的歌曲
Set fs=Server.CreateObject("Scripting.FileSystemObject")
Set fdir=fs.GetFolder(tmpfolder)
Dim theFiles( )
ReDim theFiles( 500 )
currentSlot = -1

For each thing in fdir.Files
	strExt=lcase(right(thing.Name,4))
	select case strExt
	case ".mp3"
		fname = thing.Name
		currentSlot = currentSlot + 1
		theFiles(currentSlot) = fname
	case ".err"
		fname = thing.Name
		currentSlot = currentSlot + 1
		theFiles(currentSlot) = fname
	case else
	end select
Next

fileCount = currentSlot
ReDim Preserve theFiles( currentSlot )

For i = fileCount TO 0 Step -1
	minmax = theFiles( 0 )
	minmaxSlot = 0
	For j = 1 To i
		mark = (strComp( theFiles(j), minmax, vbTextCompare ) > 0)
		If mark Then
			minmax = theFiles( j )
			minmaxSlot = j
		End If
	Next
	If minmaxSlot <> i Then
		temp = theFiles( minmaxSlot )
		theFiles( minmaxSlot ) = theFiles( i )
		theFiles( i ) = temp
	End If
Next

albumid		= albumid
singerid 	= albumid \ (10^Mul_Album)
classid 	= albumid \ (10^(Mul_Album+Mul_Singer))

set rsclass=server.createobject("adodb.recordset")
sqlclass="select * from class where classid="+cstr(classid)+"" 
rsclass.open sqlclass,conn,1,3
classname=rsclass("classname")
rsclass.close
set rsclass=nothing

set rssinger=server.createobject("adodb.recordset")
sqlsinger="select * from singer where singerid="+cstr(singerid)+"" 
rssinger.open sqlsinger,conn,1,3
singername=rssinger("singername")
rssinger.close
set rssinger=nothing

set rsalbum=server.createobject("adodb.recordset")
sqlalbum="select * from album where albumid="+cstr(albumid)+"" 
rsalbum.open sqlalbum,conn,1,3
albumname=rsalbum("albumname")
rsalbum.close
set rsalbum=nothing

For i = 0 To fileCount
	if lcase(right(theFiles(i),4)) <> ".err" then
		songfile = theFiles(i)
	else
		songfile = ""
	end if

	songname = theFiles(i)
	songname = left ( songname,len(songname)-4)
	if request("file_mode")=1 then
		songname = right( songname,len(songname)-instr(songname," - ")-2 )	'自动添加如下格式的文件：(刘德华01 - 马桶.mp3)
	else
		songname = right( songname,len(songname)-instr(songname,"_") )	'自动添加如下格式的文件：(A01_刘德华 - 马桶.mp3)
	end if
	
	'如果专辑中有歌就歌曲id+1, 若专辑中无歌曲则id=1
	sql="select * from song where albumid="+cstr(albumid)+" order by songid desc" 
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		songid=albumid*(10^Mul_Song)+1
	else
		if rs("songid")>(albumid*(10^Mul_Song)+999) then
		  	founderr=true
  			errmsg=errmsg+"<li>专辑歌曲数超过999,请整理</li>"
		else
			songid=rs("songid")+1
		end if
	end if
	rs.close
    
	sql="select * from song where (songid is null)" 
	rs.open sql,conn,1,3
	rs.addnew
		rs("songid")  = songid
		rs("songname")= songname
		if songfile<>"" then
			rs("songfile") = "songs/"&classname&"/"&singername&"/"&albumname&"/"&songfile&""
		else
			rs("songfile") = ""
		end if
		rs("albumid") = albumid
		rs("singerid")= singerid
		rs("classid") = classid
		rs("songdatetime") = date()
		rs("songlrc") = ""
		rs("songhot") = 1
	rs.update
	rs.close
	Response.Write theFiles(i)&"<br>"
Next
			%>
					</td>
				</tr>
				<tr>
					<td align="center">
						<a href="album_add.asp?step=CLASS">返回添加页面</a>
					</td>
				</tr>
			</table>
			</center></div>
		<%
		else
			response.write "由于以下的原因不能保存数据："
			response.write errmsg
			response.write "<br> <a href='#'onClick='history.back()'>返回</a>"
		end if
		%>
	<%
	case "SONG"
		albumid = request("albumid")
		singerid = albumid \ 10^Mul_Album
	%>
		<form method="POST" action="song_save.asp">
		<input type="hidden" name="albumid" size="15" value="<%=albumid%>">
		<div align="center"><center>
		<table border="0" cellspacing="5" width="500" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0">
			<tr>
				<td width="100%" colspan="2"><img src="images/button-ff4.gif" alt WIDTH="28" HEIGHT="28">  添加专辑歌曲</td>
			</tr>
			<tr>
				<td width="20%" align="right">歌曲文件：</td>
				<td width="80%">
					<input type="text" name="songfile" size="30" class="smallinput" maxlength="100" value>
				</td>
			</tr>
			<tr>
				<td width="20%" align="right">歌曲名称：</td>
				<td width="80%">
					<input type="text" name="songname" size="30">
				</td>
			</tr>
			<tr>
				<td width="20%" align="right" valign="top">歌曲评分：</td>
				<td width="80%">
					<select name="songhot" size="1">
						<option value="1">　1　</option>
						<option value="2">　2　</option>
						<option value="3">　3　</option>
						<option value="4">　4　</option>
						<option value="5">　5　</option>
					</select>
					&nbsp;&nbsp;<input type="reset" value=" 清除输入 " name="cmdcancel" style="background-color: rgb(255,255,255); border: 1px inset rgb(0,0,0)"></p>
				</td>
			</tr>
		</table>
		</center></div>
		<div align="center"><center>
			<input type="Image" src="images/button-ff.gif" tppabs title="下一步" name="Submit" WIDTH="73" HEIGHT="19">
		</center></div>
		</form>
<%end select%>

<br><br>
</body>
</html>
<%set rs=nothing%>
<!--#include file="conn_music_close.asp"-->
