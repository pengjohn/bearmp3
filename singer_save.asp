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
	errmsg=errmsg+"<li>�������ֲ���Ϊ��</li>"
end if

'�ж��Ƿ���ͬ������, ���������д�"'",������
if instr(singername,"'")=0 then	 '�жϸ��������Ƿ��д�"'"
	set rs=server.createobject("adodb.recordset")
	sql="select * from singer where singername='"&singername&"'" 
	rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			errmsg=errmsg
		else
			founderr=true
			errmsg=errmsg+"<li>����ͬ������</li>"
		end if
	rs.close
	set rs=nothing
end if

'ȡ����������
set rs=server.createobject("adodb.recordset")
sql="select * from class where classid="+cstr(classid)+"" 
rs.open sql,conn,1,3
classname=rs("classname")
rs.close
set rs=nothing

if founderr=false then
'����������и��������id+1, ���������޸��������id=1
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
    
'��Ӹ������ϵ����ݿ�
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

'���������֣����ר��������������
	set rsalbum=server.createobject("adodb.recordset")
	sqlalbum="select * from album where (albumid is null)" 
	rsalbum.open sqlalbum,conn,1,3
	rsalbum.addnew
		rsalbum("albumid")  = singerid*(10^Mul_Album)
		rsalbum("albumname")= singername+" - ��������"
		rsalbum("albumcover")= "images/songelse.gif"
		rsalbum("singerid") = singerid
		rsalbum("classid") = classid
		rsalbum("albumdatetime") = date()
		rsalbum("albuminfo") = "��������"
		rsalbum("albumgrade") = 0
	rsalbum.update
	rsalbum.close
	set rsalbum=nothing
%>
<html>
<head>
<title>ɧ������->��Ӹ��ֳɹ�</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body bgcolor="#000000">
<div align="center"><center>
<table border="1" width="700" cellpadding="2" bordercolor="#FFFFFF">
	<tr>
		<td width="100%" height="20"><p align="center"><b>��ӳɹ�</b></td>
	</tr>
	<tr>
		<td width="100%"><p align="left"><br>
		�������֣�<%=singername%><br>
		����Ŀ¼��<br>
<%
dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
tmpfolder1=initfolder&classname&"\"&singername&"\"
if not fs.FolderExists(tmpfolder1) then
	fs.CreateFolder(tmpFolder1)
	Response.Write("&nbsp;&nbsp;&nbsp;"&tmpfolder1&"<br>")
end if
%>

<%tmpfolder2=tmpfolder1&singername&" - ��������\"
if not fs.FolderExists(tmpfolder2) then
	fs.CreateFolder(tmpFolder2)
	Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;"&tmpfolder2)
end if
%>
    </p>
    <p align="center">�Ƿ������</p>
		<div align="center"><center>
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="60" style="border: 1px dotted rgb(255,255,255)"><p align="center"><a href=album_add.asp?step=ALBUM&singerid=<%=singerid%>>���ר��</a></td>
				<td width="40"></td>
				<td width="60" style="border: 1px dotted rgb(255,255,255)"><p align="center"><a href=song_add.asp?step=SONG&albumid=<%=singerid*(10^Mul_Album)%>>��ӵ���</a></td>
			</tr>
		</table>
	</td>
</tr>
</table>
</center></div>
<%else
	response.write "�������µ�ԭ���ܱ������ݣ�"
	response.write errmsg
	response.write "<br> <a href='#'onClick='history.back()'>����</a>"
end if
%>
<br><br>
</body>
</html>
<!--#include file="conn_music_close.asp"-->
