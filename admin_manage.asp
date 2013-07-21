<%if Session.Contents("KEY")<>"super" then%>
	<script language=vbscript>       
		MsgBox "你不是管理员！请返回。"
		location.href = "default.asp"
	</script> 
<%end if%>
<!--#include file="conn_music_open.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>骚熊音乐->超级管理员管理</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<script language=javascript>
function openScript(url, width, height) {
        var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}
</script>

<body topmargin="0" leftmargin="0" bgcolor="#000000" >
<div align="center"><center>
<table border="1" cellpadding="2" width="750" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
	<tr>
		<td width="10%">
			<form action="admin_user_add.asp" method="post">
			<div align="center"><center>
			<table border="0" cellPadding="0" cellSpacing="5" width="240">
				<tr>
					<td align="right" width="30%">用户名：</td>
					<td width="70%"><input class="smallInput" maxLength="30" name="UserName2" onfocus="this.select()" onmouseover="this.focus()" size="10"
					style="background-color: rgb(0,128,0); font-size: 9pt; color: rgb(0,255,0); border: 1px solid rgb(0,0,0)" value="输入用户名">
					</td>
				</tr>
				<tr>
					<td align="right" width="30%">用户权限：</td>
					<td width="70%">
						<select name="select" size="1">
							<option selected value="input">初级管理员</option>
							<option value="check">高级管理员</option>
							<option value="super">系统管理员</option>
							<option value="user">普通会员</option>
						</select>
					</td>
				</tr>
				<tr>
					<td align="right" width="30%">密码：</td>
					<td width="70%"><input class="smallInput" maxLength="30" name="Passwd2" onfocus="this.select()" onmouseover="this.focus()" size="10"
          				style="background-color: rgb(0,128,0); font-size: 9pt; color: rgb(0,255,0); border: 1px solid #000000" value="输入密码">
					</td>
				</tr>
				<tr align="center">
					<td align="right" width="30%"></td>
					<td align="center" width="70%"><div align="left"><p><input class="smallInput" name="submit" type="submit" value="增加">
						<input class="smallInput" name="Submit" type="submit" value="取消">
					</td>
				</tr>
			</table>
			</center></div>
			</form>
		</td>
		<td width="90%" align="center">
			<table width="100%" height="39" border="1" bordercolor="#FFFFFF">
				<tr>
					<td align="center"><font color="#ffffff">ID号</font></td>
					<td align="center"><font color="#ffffff">用户名</font></td>
					<td align="center"><font color="#ffffff">密码</font></td>
					<td align="center"><font color="#ffffff">权限</font></td>
					<td align="center"><font color="#ffffff">积分</font></td>
					<td align="center"><font color="#ffffff">IP</font></td>
					<td align="center"><font color="#ffffff">管理</font></td>
				</tr>
<%
dim rs
set rs=server.CreateObject("ADODB.RecordSet")  
rs.open "select * from user where userrank<>'user'",conn,1
do while NOT rs.EOF%> 
	<tr>
		<td align="center"><%=rs("userid")%></td>
		<td align="center"><A href="javascript:openScript('user_show.asp?username=<%=rs("username")%>',400,220)"><%=rs("username")%></a></td>
		<td align="center"><%=rs("userpassword")%></td>
		<td align="center"><%=rs("userrank")%></td>
		<td align="center"><%=rs("userscore")%></td>
		<td align="center"><%=rs("userip")%></td>
		<td align="center"><a href="admin_user_del.asp?id=<%=rs("userid")%>&amp;name=del">删除</a></td>
	</tr>
<%
rs.MoveNext
loop
rs.close
set rs=nothing
%>
	<tr>
		<td colspan="7">
			系统管理员： super<br>
			分类管理员： check<br>
			歌曲管理员： input<br>
			一般的会员： user<br>
		</td>
	</tr>
	</table>
    </td>
  </tr>
</table>
</center></div>
</body>
</html>
<!--#include file="conn_music_close.asp"-->
