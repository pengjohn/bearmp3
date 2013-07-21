<%
dim classid
classid = request("classid")
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>骚熊音乐->添加歌手</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body bgcolor="#000000">
<form method="POST" action="singer_save.asp">
<input type="hidden" name="classid" size="15" value="<%=classid%>">
<div align="center"><center>
<table border="1" cellspacing="0" width="80%" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0">
	<tr>
		<td width="100%" height="20"><div align="center"><center><p><b>添加歌手</b></td>
	</tr>
	<tr align="center">
		<td width="100%">
			<table border="0" cellspacing="1" width="100%">
				<tr>
					<td width="20%" align="right" height="30">歌手名字：</td>
					<td width="80%" height="30"><input type="text" name="singername" size="25"></td>
				</tr>
				<tr>
					<td width="20%" align="right" height="30">索引字母：</td>
					<td width="80%" height="30">
						<select name="singerindex" size="1">
							<option value="0" selected>　0　</option>
							<option value="A">　A　</option>
							<option value="B">　B　</option>
							<option value="C">　C　</option>
							<option value="D">　D　</option>
							<option value="E">　E　</option>
							<option value="F">　F　</option>
							<option value="G">　G　</option>
							<option value="H">　H　</option>
							<option value="I">　I　</option>
							<option value="J">　J　</option>
							<option value="K">　K　</option>
							<option value="L">　L　</option>
							<option value="M">　M　</option>
							<option value="N">　N　</option>
							<option value="O">　O　</option>
							<option value="P">　P　</option>
							<option value="Q">　Q　</option>
							<option value="R">　R　</option>
							<option value="S">　S　</option>
							<option value="T">　T　</option>
							<option value="U">　U　</option>
							<option value="V">　V　</option>
							<option value="W">　W　</option>
							<option value="X">　X　</option>
							<option value="Y">　Y　</option>
							<option value="Z">　Z　</option>
						</select>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</center></div>
<div align="center"><center><p>
  	<input type="submit" value=" 添 加 " name="cmdok"style="background-color: rgb(0,0,0); color: rgb(255,255,255); border: 1px dotted rgb(255,255,255)">&nbsp;
  	<input type="reset" value=" 清 除 " name="cmdcancel" style="background-color: rgb(0,0,0); color: rgb(255,255,255); border: 1px dotted rgb(255,255,255)"></p>
</center></div>
</form><br><br>
</body>
</html>
