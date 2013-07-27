<!--#include file="conn_music_open.asp"-->
<!--#include file="conn_define.asp"-->
<!--#include file="conn_function.asp"-->

<%
Session.Timeout=600
dim totalPut   
dim CurrentPage
dim k
dim keyword
dim sql,rs
dim classname,classid
dim singername,singerid
dim albumid,albumname
dim titlename

classid	=request("classid")
singerid	=request("singerid")
albumid	=request("albumid")
keyword	=request("keyword")			'歌曲查找
goodcomment_id = request("goodcomment_id")

if not isempty(request("page")) then
	currentPage=cint(request("page"))
else
	currentPage=1
end if

set rs=server.createobject("adodb.recordset")

'类型
if classid<>"" then
	rs.open "select * from class where classid="&cstr(classid),conn,1,1
	if not rs.eof then
		classname=rs("classname")
	end if
	rs.close
end if

'歌手
if singerid<>"" then
	rs.open "select * from singer,class where singer.classid = class.classid and singerid="&cstr(singerid),conn,1,1
	if not rs.eof then
		singername=rs("singername")
		classname=rs("classname")
	end if
	rs.close
end if

'专辑
if albumid<>"" then
	rs.open "select * from album,singer,class where album.singerid=singer.singerid and singer.classid=class.classid and albumid="&cstr(albumid),conn,1,1
	if not rs.eof then
		albumname	= rs("albumname")
		singername	= rs("singername")
		classname	= rs("classname")
		albumcover	= rs("albumcover")
		albumname	= rs("albumname")
	end if
	rs.close
end if

if Session.Contents("mp3_username")<>"" then
	username = Session.Contents("mp3_username")
	sql="select * from user where username='"&username&"'" 
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		userscore = 0			'无此user
	else
		userscore = rs("userscore")	'此user存在
		userface = rs("userimage")
	end if
	rs.close
end if
%>


<HTML>
<HEAD>
<title><%=Info_Title%><%=Info_Version%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="css/style_qvga.css">
</HEAD>

<BODY text=#000000 vLink=#003399 link=#0033cc bgColor=#fef7ed leftMargin=0 topMargin=5>
<TABLE cellSpacing=0 cellPadding=0 width=230 border=0>
	<TR>
		<TD>WePop<hr size="1"></TD>
	</TR>
	<TR>
		<TD align=middle>:<a href="default_qvga.asp">首页</a>::<a href="default_qvga.asp?classid=1">中文</a>::<a href="default_qvga.asp?classid=2">欧美</a>::<a href="default_qvga.asp?classid=3">日韩</a>::<a href="default_qvga.asp?classid=4">纯音乐::<a href=bbs.asp>论坛</a>:<hr size="1"></TD>
	</TR>

	<form action="user_chklogin.asp?loginmode=1" method="post">
	<tr>
		<td>
			<%if Session.Contents("mp3_username")<>"" then%>
				欢迎你，<%=Session.Contents("mp3_username")%>。
			<%else%>
					账号:<input type="text" name="username" maxlength="20" size="5">
					密码:<input type="password" name="password" maxlength="16" size="5">
					<input style="background-color: #CCCCCC;font-size:9pt; border: 1 solid #000000" type="submit" name="Submit" value="登录">
					<a href="user_reg.asp">注册</a>
			<%end if%>
			<hr size="1">
		</td>
	</tr>
	</form>

	<form action="default_qvga.asp?classid=81" method="post">
	<TR>
		<TD>歌曲搜索：<input type="text" name="keyword" maxlength="20" size="10">  <input style="background-color: #CCCCCC;font-size:9pt; border: 1 solid #000000" type="submit" name="Submit" value="本地搜索"><hr size="1"></TD>
	</TR>
	</form>

	<form action="http://www.pengjohn.com/search.asp" method="post">
	<TR>
		<TD>歌曲搜索：<input type="text" name="keyword" maxlength="20" size="10">  <input style="background-color: #CCCCCC;font-size:9pt; border: 1 solid #000000" type="submit" name="Submit" value="百度搜索"><hr size="1"></TD>
	</TR>
	</form>
	 
	<TR>
		<TD bgColor=#00b0c0>内容：</TD>
	</TR>
	<TR>
		<TD>
<% '歌曲类型classid从00-21, 80-99为功能扩展 
	select case classid
		case ""
			if singerid<>"" then
				sql="select * from album where singerid="+cstr(singerid)+" order by ((albumid mod 100)<>0),albumname" 
				cenbar = "Images/cenbar_listalbum.GIF"
				listalbum
			else		'singer=0
				if albumid<>"" then
					sql="select * from song where albumid="+cstr(albumid)+" order by songid" 
					cenbar = "Images/cenbar_listalbumsongs.GIF"		'专辑曲目列表
					listalbumsongs
				else		'albumid=0
					sql="select * from album where (albumid mod 100<>0) order by album_id desc"
					titlename = "新增专辑"
					cenbar = "Images/cenbar_newablum.gif"				'新增专辑"
					newalbum_list
				end if
			end if
		case 80
			sql="select * from song,album where song.albumid=album.albumid order by song_id desc"
			titlename = "新增歌曲"
			cenbar = "Images/cenbar_newsong.gif"				'新增歌曲"
			listsongs
		case 81
			sql="select * from song,album where song.albumid=album.albumid and song.songname Like '%"& keyword &"%' order by song.song_id desc"
			titlename = "歌名关键字“<font color=#00FF00>"&keyword&"</font>”歌曲"
			cenbar = "Images/cenbar_result.GIF"	'曲目关键字
			listsongs
		case 82
			sql="select * from song,album,singer where album.albumid=song.albumid and song.singerid=singer.singerid and singer.singername Like '%"& keyword &"%' order by album.albumid desc"
			titlename = "歌手关键字“<font color=#00FF00>"&keyword&"</font>”歌曲"
			cenbar = "Images/cenbar_result.GIF"	'歌手关键字
			listsongs
		case 83
			sql="select * from song,album where song.albumid=album.albumid and album.albumname Like '%"& keyword &"%' order by album.singerid,album.albumid"
			titlename = "专辑关键字“<font color=#00FF00>"&keyword&"</font>”歌曲"
			cenbar = "Images/cenbar_result.GIF"	'专辑关键字
			listsongs
		case 84
			sql="select * from song,album where song.albumid=album.albumid and songlrc Like '%"& keyword &"%' order by song_id desc"
			titlename = "歌词关键字“<font color=#00FF00>"&keyword&"</font>”歌曲"
			cenbar = "Images/cenbar_result.GIF"	'歌词关键字
			listsongs
		case 86
			sql="select * from song,album where song.albumid=album.albumid order by songhit desc"
			titlename = "下载排行"
			listsongs
		case 87
			titlename = "骚熊推荐"
			goodcomment_list
		case 88
			sql="select * from song,album where song.albumid=album.albumid and songfile='' order by song.albumid desc"
			titlename = "悬赏歌曲"
			listsongs
		case 89
			sql="select * from album where albumcover='images/nocover.gif' order by albumname" 
			titlename = "悬赏CD专辑封面"
			listalbum
		case 90
			sql="select * from song,album,singer where album.albumid=song.albumid and song.singerid=singer.singerid and singer.singername='"& keyword &"' order by album.albumid desc"
			titlename = "歌手“<font color=#00FF00>"&keyword&"</font>”所有歌曲列表"
			cenbar = "Images/cenbar_listalbumsongs.GIF"	'歌手所有歌曲列表
			listsongs
		case 91
			sql="select * from song,album where song.albumid=album.albumid and songhot>=5 order by song.albumid desc"
			titlename = "推荐歌曲"
			cenbar = "Images/cenbar_best.GIF"	'推荐歌曲
			listsongs
		case 100
			sql="select * from user order by userid desc"
			user_list
		case 0
			cenbar = "Images/cenbar_else.gif"				'其他歌曲
			sql="select * from singer where classid="+cstr(classid)+" order by singerindex, singername" 
			listsinger
		case 1
			cenbar = "Images/cenbar_cn.gif"					'华语歌手
			sql="select * from singer where classid="+cstr(classid)+" order by singerindex, singername" 
			listsinger
		case 2
			cenbar = "Images/cenbar_en.gif"					'欧美歌手
			sql="select * from singer where classid="+cstr(classid)+" order by singerindex, singername" 
			listsinger
		case 3
			cenbar = "Images/cenbar_ja_co.gif"				'日韩歌手
			sql="select * from singer where classid="+cstr(classid)+" order by singerindex, singername" 
			listsinger
		case 4
			cenbar = "Images/cenbar_NewAge.gif"				'New Age
			sql="select * from singer where classid="+cstr(classid)+" order by singerindex, singername" 
			listsinger
		case else
			sql="select * from singer where classid="+cstr(classid)+" order by singerindex, singername" 
			listsinger
	end select					
	%>
		</TD>
	</TR>
	<TR>
		<TD align=right></TD>
	</TR>
	<TR>
		<TD></TD>
	</TR>
	<TR>
		<TD><hr size="1">CopyRight(c) 2008 WePop All Rights Reserved</TD>
	</TR>
</TABLE>
</BODY>
</HTML>
<%set rs = nothing%>
<!--#include file="conn_music_close.asp"-->


<%'------------------------------------歌手列表'------------------------------------%>
<%sub listsinger%>
			<TABLE height=81 cellSpacing=0 cellPadding=3 width="97%" border=0>
	<%
	dim i 
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then 
		response.write "<p align='center'>暂时没有收集</p>" 
	end if
	%>
				<tr>
					<td width="100%" colspan="<%=MaxPerRow_Singer%>">当前位置：<%=classname%> → 歌手列表</td>
				</tr>
				<tr>
		<%
		i=1
		preindex = ""
		do while not rs.eof
		if rs("singerindex")<>preindex then
			if i<MaxPerRow_Singer then
				colspannum = MaxPerRow_Singer-i+1
				response.write "<td colspan='"& colspannum &"' width='"& (colspannum*100/MaxPerRow_Singer) &"%'></td>"
			end if
			response.write "</tr><tr><td colspan='"&MaxPerRow_Singer&"' height='1' bgcolor='#FFFFFF'></td></tr>"
			response.write "<tr><td colspan='"&MaxPerRow_Singer&"'><b>［"&rs("singerindex")&"］</b></td></tr><tr>"
			preindex = rs("singerindex")
			i=1
		end if
		response.write "<td width="&(100/MaxPerRow_Singer)&"% align=left>⊙<a href='default_qvga.asp?singerid="&rs("singerid")&"'>"&rs("singername")&"</a></td>"
		if i>=MaxPerRow_Singer then
			i=1
			response.write "</tr><tr>"
		else
			i=i+1
		end if
	  	rs.movenext
		loop
		rs.close
		response.write "</tr><tr><td colspan='"&MaxPerRow_Singer&"'><hr size=1 color=#ffffff></td></tr>"
		%>
				</tr>
			</TABLE>
<%end sub%>

<%''------------------------------------专辑列表'------------------------------------%>
<%sub listalbum%>
			<TABLE height=81 cellSpacing=0 cellPadding=3 width="97%" border=0>
			<%
			rs.open sql,conn,1,1
			if rs.eof and rs.bof then 
				response.write "<p align='center'>暂时没有收集</p>" 
			else 
			%>
				<tr>
					<td>当前位置：<a href="default_qvga.asp?classid=<%=(singerid\10^Mul_Singer)%>"><%=classname%></a> → 歌手［<%=singername%>］专辑列表</td>
				</tr>
				<tr>
					<td><a href="default_qvga.asp?classid=90&keyword=<%=singername%>">所有歌曲列表</a></td>
				</tr>
				<tr>
					<td>专辑列表：</td>
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
								<td align="left"><a href="default_qvga.asp?albumid=<%=rs("albumid")%>"><%=rs("albumname")%></a></td>
							</tr>
							<%
							rs.movenext
							loop
							%>
						</table>
					</td>
				</tr>
				<%end if
				rs.close
				%>
			</TABLE>
<%end sub%>

<%''------------------------------------专辑歌曲列表'------------------------------------%>
<% sub listalbumsongs %>
			<TABLE cellSpacing=0 cellPadding=3 width="97%" border=0>
				<tr>
					<td>当前位置：<a href="default_qvga.asp?classid=<%=(albumid\10^(Mul_Singer+Mul_album))%>"><%=classname%></a> → <a href="default_qvga.asp?singerid=<%=(albumid\10^Mul_album)%>"><%=singername%></a> → 专辑［<%=albumname%>］歌曲列表</td>
				</tr>
				<tr>
					<td width="100%" align="center">
						<%show_cd(albumcover)%>
						<br><br><%=albumname%><br><br>
					</td>
				</tr>
				<tr>
					<td>
						<table border="1" width="100%" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#CCCCCC" bordercolorlight="#000000">
							<tr align="center">
								<td width="35%" height="22" background="images/dh-di.jpg">歌曲编号</td>
								<td width="65%" background="images/dh-di.jpg">歌曲名称</td>
							</tr>
							<%
							rs.open sql,conn,1,1
							do while not rs.eof
							%>
							<tr align="center">
								<td><a href=Lyrics_dl.asp?musicname=<%=rs("songname")%>&Artist=<%=singername%>><%=rs("songid")%></a></td>
								<td>
									<%
									if rs("songfile")<>"" then
										response.write "<a href='"&rs("songfile")&"'>"&rs("songname")&"</a>"
									else
										response.write rs("songname")&"(缺)"
									end if%>
								</td>
							</tr>
							<%
							rs.movenext
							loop
							rs.close
							%>
						</table>
					</td>
				</tr>
			</TABLE>
<% end sub %>

<%''------------------------------------全歌曲列表------------------------------------%>
<%sub listsongs%>
			<TABLE height=81 cellSpacing=0 cellPadding=3 width="97%" border=0>
<%	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<p align='center'>暂时没有收集</p>"
	else
		totalPut=rs.recordcount
		if currentpage<1 then
			currentpage=1
		end if
		if (currentpage-1)*MaxPerPage_Songs>totalput then
			if (totalPut mod MaxPerPage_Songs)=0 then
				currentpage= totalPut \ MaxPerPage_Songs
			else
				currentpage= totalPut \ MaxPerPage_Songs + 1
			end if
		end if
		if currentPage=1 then
			showpage totalput,MaxPerPage_Songs,"default_qvga.asp"
			showsongs
		else
			if (currentPage-1)*MaxPerPage_Songs<totalPut then
				rs.move  (currentPage-1)*MaxPerPage_Songs
				dim bookmark
				bookmark=rs.bookmark
				showpage totalput,MaxPerPage_Songs,"default_qvga.asp"
				showsongs
			else
				currentPage=1
				showpage totalput,MaxPerPage_Songs,"default_qvga.asp"
				showsongs
			end if
		end if
	end if
	rs.close
%>
			</TABLE>
<%end sub %>

<%''------------------------------------全歌曲列表分页'------------------------------------%>
<%
sub showsongs
	dim i
	i=0
	%>
		<tr>
			<td width="100%">当前位置：<%=titlename%></td>
		</tr>
		<tr>
			<td>
				<table border="1" width="100%" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#CCCCCC" bordercolorlight="#000000">
					<tr align="center">
						<td width="35%" height="22" background="images/dh-di.jpg">歌曲名称</td>
						<td width="35%" background="images/dh-di.jpg">所属专辑</td>
					</tr>
					<%do while not rs.eof
					%>
					<tr align="center">
						<td align="left">
							<%
							if rs("songfile")<>"" then
								response.write (i+1)&". <a href='"&rs("songfile")&"'>"&rs("songname")&"</a>"
							else
								response.write (i+1)&". "&rs("songname")&"(缺)"
							end if%>
						</td>
						<td>
						<%if rs("albumid")=0 then
							response.write "会员上传"
						else
							response.write "<a href='default_qvga.asp?albumid="&rs("albumid")&"'>"&rs("albumname")&"</a>"
						end if
						%>
						</td>						
					</tr>
					<%
					i=i+1
					if i>=MaxPerPage_Songs then exit do
					rs.movenext
					loop
					%>
				</table>
			</td>
		</tr>
<% end sub %>

<%'------------------------------------分页信息'------------------------------------%>
<%
function showpage(totalnumber,maxperpage,filename)
  	dim n
  	if totalnumber mod maxperpage=0 then
     		n= totalnumber \ maxperpage
  	else
     		n= totalnumber \ maxperpage+1
  	end if
  	response.write "<form method=Post action="&filename&"?keyword="&keyword&"&classid="&classid&">"
  	response.write "<center>共<b>"&totalnumber&"</b>首｜"
  	if CurrentPage<2 then
		response.write "首页&nbsp;"
		response.write "上页&nbsp;"
  	else
		response.write "<a href="&filename&"?keyword="&keyword&"&page=1&classid="&classid&">首页</a>&nbsp;"
		response.write "<a href="&filename&"?keyword="&keyword&"&page="&CurrentPage-1&"&classid="&classid&">上页</a>&nbsp;"
  	end if
  	if n-currentpage<1 then
		response.write "下页&nbsp;"
		response.write "尾页"
  	else
		response.write "<a href="&filename&"?keyword="&keyword&"&page="&(CurrentPage+1)&"&classid="&classid&">下页</a>&nbsp;"
		response.write "<a href="&filename&"?keyword="&keyword&"&page="&n&"&classid="&classid&">尾页</a>"
  	end if
   response.write "｜<strong>"&CurrentPage&"/"&n&"</strong>页"
	%>
	转:<select name="page" size="1" style="background-color:white;FONT-SIZE: 9pt;color:black" onchange="javascript:submit()">
	<%for i = 1 to n%>   
		<option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>第<%=i%>页</option>   
	<%next%>   
	</select>
	</form>
<% end function %>


<%''------------------------------------新增专辑'------------------------------------%>
<%sub newalbum_list%>
			<TABLE cellSpacing=0 cellPadding=3 width="97%" border=0>
<%	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<p align='center'>暂时收录专辑</p>"
	else
		totalPut=rs.recordcount
		if currentpage<1 then
			currentpage=1
		end if
		if (currentpage-1)*MaxPerPage_NewAlbum>totalput then
			if (totalPut mod MaxPerPage_NewAlbum)=0 then
				currentpage= totalPut \ MaxPerPage_NewAlbum
			else
				currentpage= totalPut \ MaxPerPage_NewAlbum + 1
			end if
		end if
		if currentPage=1 then
			newalbum_showpage totalput,MaxPerPage_NewAlbum,"default_qvga.asp"
			newalbum_show
		else
			if (currentPage-1)*MaxPerPage_NewAlbum<totalPut then
				rs.move  (currentPage-1)*MaxPerPage_NewAlbum
				dim bookmark
				bookmark=rs.bookmark
				newalbum_showpage totalput,MaxPerPage_NewAlbum,"default_qvga.asp"
				newalbum_show
			else
				currentPage=1
				newalbum_showpage totalput,MaxPerPage_NewAlbum,"default_qvga.asp"
				newalbum_show
			end if
		end if
	end if
	rs.close
%>
			</TABLE>
<%end sub %>

<%'------------------------------------新增专辑分页'------------------------------------%>
<%sub newalbum_show%>
	<%
	i=0
	do while not rs.eof%>
		<TR>
			<td align="center">
			<%show_cd(rs("albumcover"))%>
			</td>
		</TR>
		<TR>
			<td align="center">
				<a href="default_qvga.asp?albumid=<%=rs("albumid")%>"><%=rs("albumname")%></a><br><br>
				入库时间：<%=rs("albumdatetime")%>
			</td>
		</TR>
		<%
		i=i+1
		if i>=MaxPerPage_NewAlbum then exit do
		rs.movenext
		loop
		%>
<% end sub %>

<%'------------------------------------新增专辑分页信息'------------------------------------%>
<%
function newalbum_showpage(totalnumber,maxperpage,filename)
  	dim n
  	if totalnumber mod maxperpage=0 then
     		n= totalnumber \ maxperpage
  	else
     		n= totalnumber \ maxperpage+1
  	end if
  	response.write "<form method=Post action="&filename&"?classid="&classid&">"
  	response.write "<tr><td>"
  	response.write "<center>专辑共<b>"&totalnumber&"</b>张｜"
  	if CurrentPage<2 then
		response.write "首页&nbsp;"
		response.write "上页&nbsp;"
  	else
		response.write "<a href="&filename&"?page=1&classid="&classid&">首页</a>&nbsp;"
		response.write "<a href="&filename&"?page="&CurrentPage-1&"&classid="&classid&">上页</a>&nbsp;"
  	end if
  	if n-currentpage<1 then
		response.write "下页&nbsp;"
		response.write "尾页"
  	else
		response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&classid="&classid&">下页</a>&nbsp;"
		response.write "<a href="&filename&"?page="&n&"&classid="&classid&">尾页</a>"
  	end if
   	response.write "｜<strong>"&CurrentPage&"/"&n&"</strong>页"
	%>
	转:<select name="page" size="1" style="background-color:white;FONT-SIZE: 9pt;color:black" onchange="javascript:submit()">
	<%for i = 1 to n%>   
		<option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>第<%=i%>页</option>   
	<%next%>   
	</select>
	</td></tr>
	</form>
<% end function %>

<%'------------------------------------显示仿CD样式'------------------------------------%>
<%function show_cd(cdcover)%>
<img src="<%=cdcover%>"  width="200" height="200" border="0">
<%end function%>
