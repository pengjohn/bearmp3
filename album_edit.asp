<!--#include file="conn_music_open.asp"-->
<%
dim album_id
dim rs,sql
dim albumid,albumname,albumcover,albumgrade

albumid = request("albumid")

set rs=server.createobject("adodb.recordset")
sql="select * from album where albumid="+cstr(albumid)+""
rs.open sql,conn,1,3
	album_id=rs("album_id")
	albumname=rs("albumname")
	albumcover=rs("albumcover")
	albumgrade=rs("albumgrade")
rs.close
set rs=nothing
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>骚熊音乐->修改专辑资料</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#000000" >
<form method="POST" action="album_edit_save.asp">
<input type="hidden" name="album_id" size="15" value="<%=album_id%>">
<div align="center"><center>
	<table border="1" cellspacing="0" width="80%" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0">
		<tr>
			<td width="100%" height="20"><div align="center"><center><p><b>专辑信息</b></td>
		</tr>
		<tr align="center">
			<td width="100%">
				<table border="0" cellspacing="1" width="100%">
					<tr>
						<td width="20%" align="right" height="30">专辑编号：</td>
						<td width="80%" height="30"><input type="text" name="albumid" size="25" value='<%=albumid%>'></td>
					</tr>
					<tr>
						<td width="20%" align="right" height="30">专辑名称：</td>
						<td width="80%" height="30"><input type="text" name="albumname" size="25" value='<%=albumname%>'></td>
					</tr>
					<tr>
						<td width="20%" align="right" height="30">专辑封面：</td>
						<td width="80%" height="30"><input type="text" name="albumcover" size="70" maxlength="200" value='<%=albumcover%>'></td>
					</tr>
					<tr>
						<td width="20%" align="right" valign="top" height="20">专辑评分：</td>
						<td width="80%">
							<select name="albumgrade" size="1">
								<option value="1">　1　</option>
								<option value="2">　2　</option>
								<option value="3">　3　</option>
								<option value="4">　4　</option>
								<option value="5">　5　</option>
							</select>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
  </center></div><div align="center"><center><p>
  	<input type="submit" value=" 添 加 " name="cmdok"style="background-color: rgb(0,0,0); color: rgb(255,255,255); border: 1px dotted rgb(255,255,255)">&nbsp;
  	<input type="reset" value=" 清 除 " name="cmdcancel" style="background-color: rgb(0,0,0); color: rgb(255,255,255); border: 1px dotted rgb(255,255,255)"></p>
  </center></div>
</form><br><br>
</body>
</html>
<!--#include file="conn_music_close.asp"-->
