<%' step: CLASS -> SINGER -> ALBUM -> SAVEALBUM(auto add song) -> SONG%>

<!--#include file="conn_music_open.asp"-->
<!--#include file="conn_define.asp"-->
<%set rs=server.createobject("adodb.recordset")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>ɧ������-&gt;���ר��</title>
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
				<td width="100%" height="20">��һ����ר����� �� ����</td>
			</tr>
			<tr align="center">
				<td width="100%">
					<img src="images/button-ff4.gif" alt WIDTH="28" HEIGHT="28">  ��ѡ���������ͣ�
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
			<input type="Image" src="images/button-ff.gif" tppabs title="��һ��" name="Submit" WIDTH="73" HEIGHT="19">
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
				<td width="100%" height="20">�ڶ�����ר����� �� ���֣�</td>
			</tr>
			<tr align="center">
				<td width="100%"><img src="images/button-ff4.gif" alt WIDTH="28" HEIGHT="28">  ��ѡ����֣�
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
					<br><br>�����ֲ����б��У���<a href="singer_add.asp?classid=<%=classid%>">��Ӹ���</a>
				</td>
			</tr>
		</table>
		</center></div>
		<div align="center"><center><p>
			<a href="#" onClick="history.back()"><img src="IMAGES/button-fb.gif" alt="��һ��" border="0" WIDTH="73" HEIGHT="19"></a>&nbsp;&nbsp;
			<input type="Image" src="images/button-ff.gif" tppabs title="��һ��" name="Submit" WIDTH="73" HEIGHT="19">
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
				<td width="100%" height="20">��������ר����� �� ר����Ϣ��</td>
			</tr>
			<tr align="center">
				<td width="100%">
					<table border="0" cellspacing="1" width="100%">
						<tr>
							<td width="20%" align="right" height="30">ר�����ƣ�</td>
							<td width="80%" height="30">
								<input type="text" name="albumname" size="25" maxlength="255">
								&nbsp;&nbsp;<input type="reset" value=" ������� " name="cmdcancel" style="background-color: rgb(255,255,255); border: 1px inset rgb(0,0,0)">
							</td>
						</tr>
						<input type="hidden" name="albumgrade" value="1">
<%
'						<tr>
'							<td width="20%" align="right" valign="top" height="20">ר�����֣�</td>
'							<td width="80%">
'								<select name="albumgrade" size="1">
'									<option value="1">��1��</option>
'									<option value="2">��2��</option>
'									<option value="3">��3��</option>
'									<option value="4">��4��</option>
'									<option value="5">��5��</option>
'								</select>
'							</td>
'						</tr>
%>
						<tr>
							<td width="20%" align="right" height="30">����ģʽ��</td>
							<td width="80%" height="30">
								<input type=radio value=1 name=file_mode checked>����01 - ����.mp3<br>
<%'								<input type=radio value=2 name=file_mode> 01_���� - ����.mp3%>
							</td>
						</tr>
						<tr>
							<td width="20%" align="right" height="30">ר�����棺</td>
							<td width="80%" height="30">�ļ���ͬר��������չ��Ϊ"jpg"��"gif"</td>
						</tr>
						<tr>
							<td width="20%" align="right" valign="top">����ר����</td>
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
			<a href="#" onClick="history.back()"><img src="IMAGES/button-fb.gif" alt="��һ��" border="0" WIDTH="73" HEIGHT="19"></a>&nbsp;&nbsp;
			<input type="Image" src="images/button-ff.gif" tppabs title="��һ��" name="Submit" WIDTH="73" HEIGHT="19">
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
			errmsg=errmsg+"<li>ר�����ֲ���Ϊ��</li>"
		end if
		'�ж��Ƿ���ͬ��ר��, ��ר�����д�"'",������
		if instr(albumname,"'")=0 then	 '�ж�ר�������Ƿ��д�"'"
			sql="select * from album where albumname='"&albumname&"'" 
			rs.open sql,conn,1,3
				if rs.eof and rs.bof then
					errmsg=errmsg
				else
					founderr=true
					errmsg=errmsg+"<li>����ͬ��ר��</li>"
				end if
			rs.close
		end if
		'ȡ����������
		sql="select * from class where classid="+cstr(classid)+"" 
		rs.open sql,conn,1,3
			classname=rs("classname")
		rs.close
		'ȡ�ø�������
		sql="select * from singer where singerid="+cstr(singerid)+"" 
		rs.open sql,conn,1,3
			singername=rs("singername")
		rs.close

		if founderr=false then
		'���������ר����ר��id+1, ��������ר����ר��id=1
			sql="select * from album where singerid="+cstr(singerid)+" order by albumid desc" 
			rs.open sql,conn,1,3
				if rs.eof and rs.bof then
					albumid=singerid*(10^Mul_Album)+1
				else
					albumid=rs("albumid")+1
				end if
			rs.close
			
			'�Զ��жϷ����ļ� *.jpg -> *.gif -> nocover.gif
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

			'���ר�����ϵ����ݿ�
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
					<td>���Ĳ���ר����� �� ���ר���ɹ���</td>
				</tr>
				<tr>
					<td width="100%">
						����ר����<%=albumname%><br>
						�� ����Ŀ¼��<br>   
							<%
							dim fs
							Set fs = CreateObject("Scripting.FileSystemObject")
							tmpfolder=Server.MapPath("./") & "\songs\"&classname&"\"&singername&"\"&albumname&"\"
						
							if not fs.FolderExists(tmpfolder) then
								fs.CreateFolder tmpFolder
								Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;"&tmpfolder&"<br>")
							end if
							Response.Write "�� ר�����棺<br><img src='"& albumcover &"'><br>"
							Response.Write "�� �Զ�������¸�����<br>"

'�Զ����ר��Ŀ¼�µĸ���
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
		songname = right( songname,len(songname)-instr(songname," - ")-2 )	'�Զ�������¸�ʽ���ļ���(���»�01 - ��Ͱ.mp3)
	else
		songname = right( songname,len(songname)-instr(songname,"_") )	'�Զ�������¸�ʽ���ļ���(A01_���»� - ��Ͱ.mp3)
	end if
	
	'���ר�����и�͸���id+1, ��ר�����޸�����id=1
	sql="select * from song where albumid="+cstr(albumid)+" order by songid desc" 
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		songid=albumid*(10^Mul_Song)+1
	else
		if rs("songid")>(albumid*(10^Mul_Song)+999) then
		  	founderr=true
  			errmsg=errmsg+"<li>ר������������999,������</li>"
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
						<a href="album_add.asp?step=CLASS">�������ҳ��</a>
					</td>
				</tr>
			</table>
			</center></div>
		<%
		else
			response.write "�������µ�ԭ���ܱ������ݣ�"
			response.write errmsg
			response.write "<br> <a href='#'onClick='history.back()'>����</a>"
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
				<td width="100%" colspan="2"><img src="images/button-ff4.gif" alt WIDTH="28" HEIGHT="28">  ���ר������</td>
			</tr>
			<tr>
				<td width="20%" align="right">�����ļ���</td>
				<td width="80%">
					<input type="text" name="songfile" size="30" class="smallinput" maxlength="100" value>
				</td>
			</tr>
			<tr>
				<td width="20%" align="right">�������ƣ�</td>
				<td width="80%">
					<input type="text" name="songname" size="30">
				</td>
			</tr>
			<tr>
				<td width="20%" align="right" valign="top">�������֣�</td>
				<td width="80%">
					<select name="songhot" size="1">
						<option value="1">��1��</option>
						<option value="2">��2��</option>
						<option value="3">��3��</option>
						<option value="4">��4��</option>
						<option value="5">��5��</option>
					</select>
					&nbsp;&nbsp;<input type="reset" value=" ������� " name="cmdcancel" style="background-color: rgb(255,255,255); border: 1px inset rgb(0,0,0)"></p>
				</td>
			</tr>
		</table>
		</center></div>
		<div align="center"><center>
			<input type="Image" src="images/button-ff.gif" tppabs title="��һ��" name="Submit" WIDTH="73" HEIGHT="19">
		</center></div>
		</form>
<%end select%>

<br><br>
</body>
</html>
<%set rs=nothing%>
<!--#include file="conn_music_close.asp"-->
