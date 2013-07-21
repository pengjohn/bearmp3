<!--#include file=conn_music_open.asp-->
<!--#include file=conn_function.asp-->
<%
dim sql,rs
dim username, userrank, useremail, useroicq, userscore, userinfo, userimage

username = Session.Contents("mp3_username")
userrank = "会员"

set rs=server.createobject("adodb.recordset")
sql="select * from user where username='"&username&"'"
rs.open sql,conn,1,3

if rs.eof and rs.bof then
	response.write "Sorry,读取数据出错！"
 	response.write "<br> <a href='#'onClick='history.back()'>返回</a>"
else
	select case rs("userrank")
		case "super"
			userrank = "超级管理员"
		case "check"
			userrank = "高级管理员"
		case "input"
			userrank = "初级管理员"
		case "user"
			userrank = "普通会员"
	end select
	
	userscore = rs("userscore")
	useremail = rs("useremail")
	useroicq = rs("useroicq")
	userinfo = rs("userinfo")
	if userinfo <>"" then 
		userinfo=userinfo
	else 
		userinfo="无"	
	end if
	userimage = rs("userimage")
%>

<title>骚熊音乐-&gt;更改资料</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="css/style.css">

<body bgcolor="#408080" topmargin="0" leftmargin="0">
<div align="center"><center>
<%if Session.Contents("mp3_username")=username then%>
<font color="#00FFFF">修改个人资料</font>
<form method="POST" action="user_save.asp">
<input type="hidden" name="username" size="15" value="<%=username%>">
<table border="0" width="53%" bordercolor="#000000" cellpadding="0">
	<tr>
		<td width="30%" align="right">Email: </td>
		<td width="70%"><input type="text" name="useremail" size="25" style="border: 1px inset rgb(0,0,0)" value="<%=useremail%>"></td>
	</tr>
	<tr>
		<td width="30%" align="right">OICQ:</td>
		<td width="70%"><input type="text" name="useroicq" size="20" style="border: 1px inset rgb(0,0,0)" value="<%=useroicq%>"></td>
	</tr>
	<tr>
		<td width="30%" align="right" valign="top">签名：</td>
		<td width="70%"><textarea rows="3" name="userinfo" cols="20" style="border: 1px inset rgb(0,0,0)"><%=userinfo%></textarea></td>
	</tr>
	<tr>
		<td align="right" valign="top">头像：</td>
		<td>
			<img id="face" src="<%=userimage%>" alt="个人形象代表" height="32" width="32">
			<select name="userimage" size="1" onChange="document.images['face'].src=options[selectedIndex].value;" style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">
			<option value="<%=userimage%>">--</option>
			<option value="pic/Image001.gif">Image001</option>
			<option value="pic/Image002.gif">Image002</option>
			<option value="pic/Image003.gif">Image003</option>
			<option value="pic/Image003.gif">Image003</option>
			<option value="pic/Image004.gif">Image004</option>
			<option value="pic/Image005.gif">Image005</option>
			<option value="pic/Image006.gif">Image006</option>
			<option value="pic/Image007.gif">Image007</option>
			<option value="pic/Image008.gif">Image008</option>
			<option value="pic/Image009.gif">Image009</option>
			<option value="pic/Image010.gif">Image010</option>
			<option value="pic/Image011.gif">Image011</option>
			<option value="pic/Image012.gif">Image012</option>
			<option value="pic/Image013.gif">Image013</option>
			<option value="pic/Image014.gif">Image014</option>
			<option value="pic/Image015.gif">Image015</option>
			<option value="pic/Image016.gif">Image016</option>
			<option value="pic/Image017.gif">Image017</option>
			<option value="pic/Image018.gif">Image018</option>
			<option value="pic/Image019.gif">Image019</option>
			<option value="pic/Image020.gif">Image020</option>
			<option value="pic/Image021.gif">Image021</option>
			<option value="pic/Image022.gif">Image022</option>
			<option value="pic/Image023.gif">Image023</option>
			<option value="pic/Image024.gif">Image024</option>
			<option value="pic/Image025.gif">Image025</option>
			<option value="pic/Image026.gif">Image026</option>
			<option value="pic/Image027.gif">Image027</option>
			<option value="pic/Image028.gif">Image028</option>
			<option value="pic/Image029.gif">Image029</option>
			<option value="pic/Image030.gif">Image030</option>
			<option value="pic/Image031.gif">Image031</option>
			<option value="pic/Image032.gif">Image032</option>
			<option value="pic/Image033.gif">Image033</option>
			<option value="pic/Image034.gif">Image034</option>
			<option value="pic/Image035.gif">Image035</option>
			<option value="pic/Image036.gif">Image036</option>
			<option value="pic/Image037.gif">Image037</option>
			<option value="pic/Image038.gif">Image038</option>
			<option value="pic/Image039.gif">Image039</option>
			<option value="pic/Image040.gif">Image040</option>
			<option value="pic/Image041.gif">Image041</option>
			<option value="pic/Image042.gif">Image042</option>
			<option value="pic/Image043.gif">Image043</option>
			<option value="pic/Image044.gif">Image044</option>
			<option value="pic/Image045.gif">Image045</option>
			<option value="pic/Image046.gif">Image046</option>
			<option value="pic/Image047.gif">Image047</option>
			<option value="pic/Image048.gif">Image048</option>
			<option value="pic/Image049.gif">Image049</option>
			<option value="pic/Image050.gif">Image050</option>
			<option value="pic/Image051.gif">Image051</option>
			<option value="pic/Image052.gif">Image052</option>
			<option value="pic/Image053.gif">Image053</option>
			<option value="pic/Image054.gif">Image054</option>
			<option value="pic/Image055.gif">Image055</option>
			<option value="pic/Image056.gif">Image056</option>
			<option value="pic/Image057.gif">Image057</option>
			<option value="pic/Image058.gif">Image058</option>
			<option value="pic/Image059.gif">Image059</option>
			<option value="pic/Image060.gif">Image060</option>
			<option value="pic/Image000.gif">骚熊专用</option>
			</select>  
		</td>
	</tr>
	<tr>
		<td colspan="2" align="center">
			<br><input type="Image" src="images/button-ok.gif" tppabs title="修改资料" name="Submit" align="absmiddle" border="0" WIDTH="73" HEIGHT="19">
		</td>
	</tr>
</table>
</form>
</center></div>
</body>
<%
end if
end if
rs.close
set rs=nothing
%>
<!--#include file=conn_music_close.asp-->
