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
keyword	=request("keyword")			'��������
goodcomment_id = request("goodcomment_id")

if not isempty(request("page")) then
	currentPage=cint(request("page"))
else
	currentPage=1
end if

set rs=server.createobject("adodb.recordset")

'����
if classid<>"" then
	rs.open "select * from class where classid="&cstr(classid),conn,1,1
	if not rs.eof then
		classname=rs("classname")
	end if
	rs.close
end if

'����
if singerid<>"" then
	rs.open "select * from singer,class where singer.classid = class.classid and singerid="&cstr(singerid),conn,1,1
	if not rs.eof then
		singername=rs("singername")
		classname=rs("classname")
	end if
	rs.close
end if

'ר��
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
		userscore = 0			'�޴�user
	else
		userscore = rs("userscore")	'��user����
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
		<TD align=middle>:<a href="default_qvga.asp">��ҳ</a>::<a href="default_qvga.asp?classid=1">����</a>::<a href="default_qvga.asp?classid=2">ŷ��</a>::<a href="default_qvga.asp?classid=3">�պ�</a>::<a href="default_qvga.asp?classid=4">������::<a href=bbs.asp>��̳</a>:<hr size="1"></TD>
	</TR>

	<form action="user_chklogin.asp?loginmode=1" method="post">
	<tr>
		<td>
			<%if Session.Contents("mp3_username")<>"" then%>
				��ӭ�㣬<%=Session.Contents("mp3_username")%>��
			<%else%>
					�˺�:<input type="text" name="username" maxlength="20" size="5">
					����:<input type="password" name="password" maxlength="16" size="5">
					<input style="background-color: #CCCCCC;font-size:9pt; border: 1 solid #000000" type="submit" name="Submit" value="��¼">
					<a href="user_reg.asp">ע��</a>
			<%end if%>
			<hr size="1">
		</td>
	</tr>
	</form>

	<form action="default_qvga.asp?classid=81" method="post">
	<TR>
		<TD>����������<input type="text" name="keyword" maxlength="20" size="10">  <input style="background-color: #CCCCCC;font-size:9pt; border: 1 solid #000000" type="submit" name="Submit" value="��������"><hr size="1"></TD>
	</TR>
	</form>

	<form action="http://www.pengjohn.com/search.asp" method="post">
	<TR>
		<TD>����������<input type="text" name="keyword" maxlength="20" size="10">  <input style="background-color: #CCCCCC;font-size:9pt; border: 1 solid #000000" type="submit" name="Submit" value="�ٶ�����"><hr size="1"></TD>
	</TR>
	</form>
	 
	<TR>
		<TD bgColor=#00b0c0>���ݣ�</TD>
	</TR>
	<TR>
		<TD>
<% '��������classid��00-21, 80-99Ϊ������չ 
	select case classid
		case ""
			if singerid<>"" then
				sql="select * from album where singerid="+cstr(singerid)+" order by ((albumid mod 100)<>0),albumname" 
				cenbar = "Images/cenbar_listalbum.GIF"
				listalbum
			else		'singer=0
				if albumid<>"" then
					sql="select * from song where albumid="+cstr(albumid)+" order by songid" 
					cenbar = "Images/cenbar_listalbumsongs.GIF"		'ר����Ŀ�б�
					listalbumsongs
				else		'albumid=0
					sql="select * from album where (albumid mod 100<>0) order by album_id desc"
					titlename = "����ר��"
					cenbar = "Images/cenbar_newablum.gif"				'����ר��"
					newalbum_list
				end if
			end if
		case 80
			sql="select * from song,album where song.albumid=album.albumid order by song_id desc"
			titlename = "��������"
			cenbar = "Images/cenbar_newsong.gif"				'��������"
			listsongs
		case 81
			sql="select * from song,album where song.albumid=album.albumid and song.songname Like '%"& keyword &"%' order by song.song_id desc"
			titlename = "�����ؼ��֡�<font color=#00FF00>"&keyword&"</font>������"
			cenbar = "Images/cenbar_result.GIF"	'��Ŀ�ؼ���
			listsongs
		case 82
			sql="select * from song,album,singer where album.albumid=song.albumid and song.singerid=singer.singerid and singer.singername Like '%"& keyword &"%' order by album.albumid desc"
			titlename = "���ֹؼ��֡�<font color=#00FF00>"&keyword&"</font>������"
			cenbar = "Images/cenbar_result.GIF"	'���ֹؼ���
			listsongs
		case 83
			sql="select * from song,album where song.albumid=album.albumid and album.albumname Like '%"& keyword &"%' order by album.singerid,album.albumid"
			titlename = "ר���ؼ��֡�<font color=#00FF00>"&keyword&"</font>������"
			cenbar = "Images/cenbar_result.GIF"	'ר���ؼ���
			listsongs
		case 84
			sql="select * from song,album where song.albumid=album.albumid and songlrc Like '%"& keyword &"%' order by song_id desc"
			titlename = "��ʹؼ��֡�<font color=#00FF00>"&keyword&"</font>������"
			cenbar = "Images/cenbar_result.GIF"	'��ʹؼ���
			listsongs
		case 86
			sql="select * from song,album where song.albumid=album.albumid order by songhit desc"
			titlename = "��������"
			listsongs
		case 87
			titlename = "ɧ���Ƽ�"
			goodcomment_list
		case 88
			sql="select * from song,album where song.albumid=album.albumid and songfile='' order by song.albumid desc"
			titlename = "���͸���"
			listsongs
		case 89
			sql="select * from album where albumcover='images/nocover.gif' order by albumname" 
			titlename = "����CDר������"
			listalbum
		case 90
			sql="select * from song,album,singer where album.albumid=song.albumid and song.singerid=singer.singerid and singer.singername='"& keyword &"' order by album.albumid desc"
			titlename = "���֡�<font color=#00FF00>"&keyword&"</font>�����и����б�"
			cenbar = "Images/cenbar_listalbumsongs.GIF"	'�������и����б�
			listsongs
		case 91
			sql="select * from song,album where song.albumid=album.albumid and songhot>=5 order by song.albumid desc"
			titlename = "�Ƽ�����"
			cenbar = "Images/cenbar_best.GIF"	'�Ƽ�����
			listsongs
		case 100
			sql="select * from user order by userid desc"
			user_list
		case 0
			cenbar = "Images/cenbar_else.gif"				'��������
			sql="select * from singer where classid="+cstr(classid)+" order by singerindex, singername" 
			listsinger
		case 1
			cenbar = "Images/cenbar_cn.gif"					'�������
			sql="select * from singer where classid="+cstr(classid)+" order by singerindex, singername" 
			listsinger
		case 2
			cenbar = "Images/cenbar_en.gif"					'ŷ������
			sql="select * from singer where classid="+cstr(classid)+" order by singerindex, singername" 
			listsinger
		case 3
			cenbar = "Images/cenbar_ja_co.gif"				'�պ�����
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


<%'------------------------------------�����б�'------------------------------------%>
<%sub listsinger%>
			<TABLE height=81 cellSpacing=0 cellPadding=3 width="97%" border=0>
	<%
	dim i 
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then 
		response.write "<p align='center'>��ʱû���ռ�</p>" 
	end if
	%>
				<tr>
					<td width="100%" colspan="<%=MaxPerRow_Singer%>">��ǰλ�ã�<%=classname%> �� �����б�</td>
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
			response.write "<tr><td colspan='"&MaxPerRow_Singer&"'><b>��"&rs("singerindex")&"��</b></td></tr><tr>"
			preindex = rs("singerindex")
			i=1
		end if
		response.write "<td width="&(100/MaxPerRow_Singer)&"% align=left>��<a href='default_qvga.asp?singerid="&rs("singerid")&"'>"&rs("singername")&"</a></td>"
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

<%''------------------------------------ר���б�'------------------------------------%>
<%sub listalbum%>
			<TABLE height=81 cellSpacing=0 cellPadding=3 width="97%" border=0>
			<%
			rs.open sql,conn,1,1
			if rs.eof and rs.bof then 
				response.write "<p align='center'>��ʱû���ռ�</p>" 
			else 
			%>
				<tr>
					<td>��ǰλ�ã�<a href="default_qvga.asp?classid=<%=(singerid\10^Mul_Singer)%>"><%=classname%></a> �� ���֣�<%=singername%>��ר���б�</td>
				</tr>
				<tr>
					<td><a href="default_qvga.asp?classid=90&keyword=<%=singername%>">���и����б�</a></td>
				</tr>
				<tr>
					<td>ר���б�</td>
				</tr>
				<tr>
					<td colspan="2">
						<table border="1" width="100%" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#CCCCCC" bordercolorlight="#000000">
							<tr align="center">
								<td width="12%" height="22" background="images/dh-di.jpg">ר�����</td>
								<td width="40%" background="images/dh-di.jpg">CDר��</td>
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

<%''------------------------------------ר�������б�'------------------------------------%>
<% sub listalbumsongs %>
			<TABLE cellSpacing=0 cellPadding=3 width="97%" border=0>
				<tr>
					<td>��ǰλ�ã�<a href="default_qvga.asp?classid=<%=(albumid\10^(Mul_Singer+Mul_album))%>"><%=classname%></a> �� <a href="default_qvga.asp?singerid=<%=(albumid\10^Mul_album)%>"><%=singername%></a> �� ר����<%=albumname%>�ݸ����б�</td>
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
								<td width="35%" height="22" background="images/dh-di.jpg">�������</td>
								<td width="65%" background="images/dh-di.jpg">��������</td>
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
										response.write rs("songname")&"(ȱ)"
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

<%''------------------------------------ȫ�����б�------------------------------------%>
<%sub listsongs%>
			<TABLE height=81 cellSpacing=0 cellPadding=3 width="97%" border=0>
<%	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<p align='center'>��ʱû���ռ�</p>"
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

<%''------------------------------------ȫ�����б��ҳ'------------------------------------%>
<%
sub showsongs
	dim i
	i=0
	%>
		<tr>
			<td width="100%">��ǰλ�ã�<%=titlename%></td>
		</tr>
		<tr>
			<td>
				<table border="1" width="100%" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#CCCCCC" bordercolorlight="#000000">
					<tr align="center">
						<td width="35%" height="22" background="images/dh-di.jpg">��������</td>
						<td width="35%" background="images/dh-di.jpg">����ר��</td>
					</tr>
					<%do while not rs.eof
					%>
					<tr align="center">
						<td align="left">
							<%
							if rs("songfile")<>"" then
								response.write (i+1)&". <a href='"&rs("songfile")&"'>"&rs("songname")&"</a>"
							else
								response.write (i+1)&". "&rs("songname")&"(ȱ)"
							end if%>
						</td>
						<td>
						<%if rs("albumid")=0 then
							response.write "��Ա�ϴ�"
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

<%'------------------------------------��ҳ��Ϣ'------------------------------------%>
<%
function showpage(totalnumber,maxperpage,filename)
  	dim n
  	if totalnumber mod maxperpage=0 then
     		n= totalnumber \ maxperpage
  	else
     		n= totalnumber \ maxperpage+1
  	end if
  	response.write "<form method=Post action="&filename&"?keyword="&keyword&"&classid="&classid&">"
  	response.write "<center>��<b>"&totalnumber&"</b>�ף�"
  	if CurrentPage<2 then
		response.write "��ҳ&nbsp;"
		response.write "��ҳ&nbsp;"
  	else
		response.write "<a href="&filename&"?keyword="&keyword&"&page=1&classid="&classid&">��ҳ</a>&nbsp;"
		response.write "<a href="&filename&"?keyword="&keyword&"&page="&CurrentPage-1&"&classid="&classid&">��ҳ</a>&nbsp;"
  	end if
  	if n-currentpage<1 then
		response.write "��ҳ&nbsp;"
		response.write "βҳ"
  	else
		response.write "<a href="&filename&"?keyword="&keyword&"&page="&(CurrentPage+1)&"&classid="&classid&">��ҳ</a>&nbsp;"
		response.write "<a href="&filename&"?keyword="&keyword&"&page="&n&"&classid="&classid&">βҳ</a>"
  	end if
   response.write "��<strong>"&CurrentPage&"/"&n&"</strong>ҳ"
	%>
	ת:<select name="page" size="1" style="background-color:white;FONT-SIZE: 9pt;color:black" onchange="javascript:submit()">
	<%for i = 1 to n%>   
		<option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>��<%=i%>ҳ</option>   
	<%next%>   
	</select>
	</form>
<% end function %>


<%''------------------------------------����ר��'------------------------------------%>
<%sub newalbum_list%>
			<TABLE cellSpacing=0 cellPadding=3 width="97%" border=0>
<%	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<p align='center'>��ʱ��¼ר��</p>"
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

<%'------------------------------------����ר����ҳ'------------------------------------%>
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
				���ʱ�䣺<%=rs("albumdatetime")%>
			</td>
		</TR>
		<%
		i=i+1
		if i>=MaxPerPage_NewAlbum then exit do
		rs.movenext
		loop
		%>
<% end sub %>

<%'------------------------------------����ר����ҳ��Ϣ'------------------------------------%>
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
  	response.write "<center>ר����<b>"&totalnumber&"</b>�ţ�"
  	if CurrentPage<2 then
		response.write "��ҳ&nbsp;"
		response.write "��ҳ&nbsp;"
  	else
		response.write "<a href="&filename&"?page=1&classid="&classid&">��ҳ</a>&nbsp;"
		response.write "<a href="&filename&"?page="&CurrentPage-1&"&classid="&classid&">��ҳ</a>&nbsp;"
  	end if
  	if n-currentpage<1 then
		response.write "��ҳ&nbsp;"
		response.write "βҳ"
  	else
		response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&classid="&classid&">��ҳ</a>&nbsp;"
		response.write "<a href="&filename&"?page="&n&"&classid="&classid&">βҳ</a>"
  	end if
   	response.write "��<strong>"&CurrentPage&"/"&n&"</strong>ҳ"
	%>
	ת:<select name="page" size="1" style="background-color:white;FONT-SIZE: 9pt;color:black" onchange="javascript:submit()">
	<%for i = 1 to n%>   
		<option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>��<%=i%>ҳ</option>   
	<%next%>   
	</select>
	</td></tr>
	</form>
<% end function %>

<%'------------------------------------��ʾ��CD��ʽ'------------------------------------%>
<%function show_cd(cdcover)%>
<img src="<%=cdcover%>"  width="200" height="200" border="0">
<%end function%>
