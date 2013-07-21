<%if Session.Contents("IsAdmin")<>true then%>
	<script language=vbscript>       
		MsgBox "你不是管理员。请返回"
		location.href = "javascript:history.back()"       
	</script> 
<%end if%>

<!--#include file="conn_music_open.asp"-->
<!--#include file="conn_define.asp"-->
<%
albumid = request("albumid")
conn.execute "delete * from song where albumid=" & albumid

albumid	= albumid
singerid = albumid \ (10^Mul_Album)
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

dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
tmpfolder=Server.MapPath("./") & "\songs\"&classname&"\"&singername&"\"&albumname&"\"

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

For i = 0 To fileCount
	if lcase(right(theFiles(i),4)) <> ".err" then
		songfile = theFiles(i)
	else
		songfile = ""
	end if

	songname = theFiles(i)
	songname = left ( songname,len(songname)-4)
	songname = right( songname,len(songname)-instr(songname," - ")-2 )
'自动添加如下格式的文件：(A01_刘德华 - 马桶.mp3)		
'	songname = right( songname,len(songname)-instr(songname,"_") )
	
	'如果专辑中有歌就歌曲id+1, 若专辑中无歌曲则id=1
	set rs=server.createobject("adodb.recordset")
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
	set rs=nothing
    
	set rs=server.createobject("adodb.recordset")
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
	set rs=nothing
	Response.Write theFiles(i)&"<br>"
Next

response.write "更新专辑数据成功。<br>"
%>
<script language=vbscript>       
	MsgBox "更新专辑数据成功。请返回"
	location.href = "default.asp?albumid=<%=albumid%>"
</script> 
<!--#include file="conn_music_close.asp"-->
