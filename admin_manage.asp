<%if Session.Contents("KEY")<>"super" then%>
	<script language=vbscript>       
		MsgBox "�㲻�ǹ���Ա���뷵�ء�"
		location.href = "default.asp"
	</script> 
<%end if%>
<!--#include file="conn_music_open.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ɧ������->��������Ա����</title>
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
					<td align="right" width="30%">�û�����</td>
					<td width="70%"><input class="smallInput" maxLength="30" name="UserName2" onfocus="this.select()" onmouseover="this.focus()" size="10"
					style="background-color: rgb(0,128,0); font-size: 9pt; color: rgb(0,255,0); border: 1px solid rgb(0,0,0)" value="�����û���">
					</td>
				</tr>
				<tr>
					<td align="right" width="30%">�û�Ȩ�ޣ�</td>
					<td width="70%">
						<select name="select" size="1">
							<option selected value="input">��������Ա</option>
							<option value="check">�߼�����Ա</option>
							<option value="super">ϵͳ����Ա</option>
							<option value="user">��ͨ��Ա</option>
						</select>
					</td>
				</tr>
				<tr>
					<td align="right" width="30%">���룺</td>
					<td width="70%"><input class="smallInput" maxLength="30" name="Passwd2" onfocus="this.select()" onmouseover="this.focus()" size="10"
          				style="background-color: rgb(0,128,0); font-size: 9pt; color: rgb(0,255,0); border: 1px solid #000000" value="��������">
					</td>
				</tr>
				<tr align="center">
					<td align="right" width="30%"></td>
					<td align="center" width="70%"><div align="left"><p><input class="smallInput" name="submit" type="submit" value="����">
						<input class="smallInput" name="Submit" type="submit" value="ȡ��">
					</td>
				</tr>
			</table>
			</center></div>
			</form>
		</td>
		<td width="90%" align="center">
			<table width="100%" height="39" border="1" bordercolor="#FFFFFF">
				<tr>
					<td align="center"><font color="#ffffff">ID��</font></td>
					<td align="center"><font color="#ffffff">�û���</font></td>
					<td align="center"><font color="#ffffff">����</font></td>
					<td align="center"><font color="#ffffff">Ȩ��</font></td>
					<td align="center"><font color="#ffffff">����</font></td>
					<td align="center"><font color="#ffffff">IP</font></td>
					<td align="center"><font color="#ffffff">����</font></td>
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
		<td align="center"><a href="admin_user_del.asp?id=<%=rs("userid")%>&amp;name=del">ɾ��</a></td>
	</tr>
<%
rs.MoveNext
loop
rs.close
set rs=nothing
%>
	<tr>
		<td colspan="7">
			ϵͳ����Ա�� super<br>
			�������Ա�� check<br>
			��������Ա�� input<br>
			һ��Ļ�Ա�� user<br>
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
