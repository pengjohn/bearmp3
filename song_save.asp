<!--#include file="conn_music_open.asp"-->
<!--#include file="conn_define.asp"-->

<%
dim songid , songname, songfile
dim albumid , singerid
dim classid
dim songdatetime
dim songhit
dim songlrc
dim songhot

dim classname
dim singername
dim albumname

dim rs
dim sql

dim rsalbum
dim sqlalbum

dim rssinger
dim sqlsinger

dim rsclass
dim sqlclass


songname = request("songname")
songfile = request("songfile")
songlrc  = request("songlrc")
songhot  = request("songhot")
albumid = request("albumid")
singerid = albumid \ (10^Mul_Album)
classid =  albumid \ (10^(Mul_Album+Mul_Singer))

founerr=false

if trim(request.form("songname"))="" then
	if trim(request.form("songfile"))<>"" then
		songname = left ( songfile,len(songfile)-4)
		songname = right( songname,len(songname)-instr(songname," - ")-2 )	
	else
	  founderr=true
	  errmsg=errmsg+"<li>�����ļ�����ʾ���Ʋ���ͬʱΪ��</li>"
	end if
end if

'���ר�����и�͸���id+1, ��ר�����޸�����id=1
	set rs=server.createobject("adodb.recordset")
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
	set rs=nothing
    
if founderr=false then

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

	songfile="songs/"&classname&"/"&singername&"/"&albumname&"/"&songfile&""
	songhot=request("songhot")
	songlrc=request("songlrc")
	    
	set rs=server.createobject("adodb.recordset")
	sql="select * from song where (songid is null)" 
	rs.open sql,conn,1,3
	rs.addnew
		rs("songid")  = songid
		rs("songname")= songname
		if request("songfile")<>"" then
			rs("songfile") = songfile
		else
			rs("songfile") = ""
		end if
		rs("albumid") = albumid
		rs("singerid")= singerid
		rs("classid") = classid
		rs("songdatetime") = date()
		rs("songlrc") = songlrc
		rs("songhot") = songhot
	rs.update
	rs.close
	set rs=nothing
%>
<html>
<head>
<title>ɧ������-&gt;��Ӹ����ɹ�</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#000000">
<div align="center"><center>
<table border="0" cellspacing="10" width="500" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0">
	<tr>
		<td width="100%" height="20">��Ӹ����ɹ�</td>
	</tr>
	<tr>
		<td width="100%">
		�����ļ���<%=songfile%><br>
		�������ƣ�<%=songname%>
		</td>
	</tr>
	<tr>
		<td align="center"><img src="images/button-ff4.gif" alt WIDTH="28" HEIGHT="28">  �Ƿ������ӣ�<br><br>
			<a href="#" onClick="history.back()"><img src="IMAGES/button-ok.gif" alt="��һ��" border="0" WIDTH="73" HEIGHT="19"></a>
			<a href="default.asp"><img src="IMAGES/button-cancel.gif" alt="������ҳ" border="0" WIDTH="73" HEIGHT="19"></a>		
		</td>
	</tr>
</table>
<%
else
	response.write "�������µ�ԭ���ܱ������ݣ�"
	response.write errmsg
	response.write "<br> <a href='#'onClick='history.back()'>����</a>"
end if
%>
<br><br>
</body>
</html>
<!--#include file="conn_music_close.asp"-->