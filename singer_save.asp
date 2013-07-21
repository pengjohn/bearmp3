<!--#include file="conn_music_open.asp"-->
<!--#include file="conn_define.asp"-->
<%
dim singerid
dim singername
dim singerindex
dim classid
dim singerinfo
dim classname

dim rs
dim sql

dim rsalbum
dim sqlalbum
dim tmpfolder1
dim tmpfolder2

initfolder=Server.MapPath("./")
initfolder=initfolder&"\songs\"

singername = request("singername")
singerinfo = request("singerinfo")
singerindex= request("singerindex")
classid = request("classid")

founerr=false
if trim(request.form("singername"))="" then
	founderr=true
	errmsg=errmsg+"<li>歌手名字不能为空</li>"
end if

'判断是否有同名歌手, 若歌手名中带"'",则跳过
if instr(singername,"'")=0 then	 '判断歌手名中是否有带"'"
	set rs=server.createobject("adodb.recordset")
	sql="select * from singer where singername='"&singername&"'" 
	rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			errmsg=errmsg
		else
			founderr=true
			errmsg=errmsg+"<li>已有同名歌手</li>"
		end if
	rs.close
	set rs=nothing
end if

'取得类型名字
set rs=server.createobject("adodb.recordset")
sql="select * from class where classid="+cstr(classid)+"" 
rs.open sql,conn,1,3
classname=rs("classname")
rs.close
set rs=nothing

if founderr=false then
'如果类型中有歌手则歌手id+1, 若类型中无歌手则歌手id=1
    set rs=server.createobject("adodb.recordset")
    sql="select * from singer where classid="+cstr(classid)+" order by singerid desc" 
    rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		singerid=classid*(10^Mul_Singer)+1
	else
		singerid=rs("singerid")+1
	end if
    rs.close
    set rs=nothing
    
'添加歌手资料到数据库
	set rs=server.createobject("adodb.recordset")
	sql="select * from singer where (singerid is null)" 
	rs.open sql,conn,1,3
	rs.addnew
		rs("singerid")  = singerid
		rs("singername")= singername
		rs("singerindex")= singerindex
		rs("classid") = classid
		rs("singerinfo") = singerinfo
	rs.update
	rs.close
	set rs=nothing

'给新增歌手，添加专辑“其他单曲”
	set rsalbum=server.createobject("adodb.recordset")
	sqlalbum="select * from album where (albumid is null)" 
	rsalbum.open sqlalbum,conn,1,3
	rsalbum.addnew
		rsalbum("albumid")  = singerid*(10^Mul_Album)
		rsalbum("albumname")= singername+" - 其他单曲"
		rsalbum("albumcover")= "images/songelse.gif"
		rsalbum("singerid") = singerid
		rsalbum("classid") = classid
		rsalbum("albumdatetime") = date()
		rsalbum("albuminfo") = "其他单曲"
		rsalbum("albumgrade") = 0
	rsalbum.update
	rsalbum.close
	set rsalbum=nothing
%>
<html>
<head>
<title>骚熊音乐->添加歌手成功</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body bgcolor="#000000">
<div align="center"><center>
<table border="1" width="700" cellpadding="2" bordercolor="#FFFFFF">
	<tr>
		<td width="100%" height="20"><p align="center"><b>添加成功</b></td>
	</tr>
	<tr>
		<td width="100%"><p align="left"><br>
		新增歌手：<%=singername%><br>
		新增目录：<br>
<%
dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
tmpfolder1=initfolder&classname&"\"&singername&"\"
if not fs.FolderExists(tmpfolder1) then
	fs.CreateFolder(tmpFolder1)
	Response.Write("&nbsp;&nbsp;&nbsp;"&tmpfolder1&"<br>")
end if
%>

<%tmpfolder2=tmpfolder1&singername&" - 其他单曲\"
if not fs.FolderExists(tmpfolder2) then
	fs.CreateFolder(tmpFolder2)
	Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;"&tmpfolder2)
end if
%>
    </p>
    <p align="center">是否继续？</p>
		<div align="center"><center>
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="60" style="border: 1px dotted rgb(255,255,255)"><p align="center"><a href=album_add.asp?step=ALBUM&singerid=<%=singerid%>>添加专辑</a></td>
				<td width="40"></td>
				<td width="60" style="border: 1px dotted rgb(255,255,255)"><p align="center"><a href=song_add.asp?step=SONG&albumid=<%=singerid*(10^Mul_Album)%>>添加单曲</a></td>
			</tr>
		</table>
	</td>
</tr>
</table>
</center></div>
<%else
	response.write "由于以下的原因不能保存数据："
	response.write errmsg
	response.write "<br> <a href='#'onClick='history.back()'>返回</a>"
end if
%>
<br><br>
</body>
</html>
<!--#include file="conn_music_close.asp"-->
