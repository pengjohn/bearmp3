<%' step: CLASS -> SINGER -> SONG%>
<!--#include file="conn_music_open.asp"-->
<!--#include file="conn_define.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>ɧ������-&gt;��Ӹ���</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#000000">
<%select case request("step")%>
	<%case "CLASS"%>
		<form method="POST" action="song_add.asp?step=SINGER">
		<div align="center"><center>
		<table border="0" cellspacing="10" width="500" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0">
			<tr>
				<td width="100%" height="20">��һ����������� �� ����</b></td>
			</tr>
			<tr align="center">
				<td width="100%">
					<table border="0" cellspacing="1" width="100%">
						<tr>
							<td width="30%" align="right" valign="top" height="20">��ѡ���������ͣ�</td>
							<td width="40%">
							<select name="classid" size="1">
								<%
								set rs=server.createobject("adodb.recordset")
								sql="select * from class"
								rs.open sql,conn,1,1
								Do until rs.eof
								%>
								<option value="<%=rs("classid")%>">= <%=rs("classname")%> =</option>
								<%rs.MoveNext
								Loop
								rs.close
								%>
							</select>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		</center></div>
		<div align="center"><center><p>
			<input type="Image" src="images/button-ff.gif" tppabs title="��һ��" name="Submit" WIDTH="73" HEIGHT="19">
		</p></center></div>
		</form>
	<%case "SINGER"
		classid = request("classid")
	%>
		<form method="POST" action="song_add.asp?step=SONG">
		<div align="center"><center>
		<table border="0" cellspacing="10" width="500" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0">
			<tr>
				<td width="100%" height="20">�ڶ������������ �� ����</b></td>
			</tr>
			<tr align="center">
				<td height="20">��ѡ����֣�
					<select name="albumid" size="1">
						<%
						set rs=server.createobject("adodb.recordset")
						sql="select * from singer where classid="+cstr(classid)+" order by singername desc"
						rs.open sql,conn,1,1
						Do until rs.eof
						%>
							<option value="<%=rs("singerid")*(10^Mul_Album)%>">= <%=rs("singername")%> =</option>
						<%rs.MoveNext
						Loop
						rs.close
						%>
					</select>
					<br><br>�����ֲ����б��У���<a href="singer_add.asp?classid=<%=classid%>">��Ӹ���</a>
				</td>
			</tr>
		</table>
		</center></div>
		<div align="center"><center>
			<a href="#" onClick="history.back()"><img src="IMAGES/button-fb.gif" alt="��һ��" border="0" WIDTH="73" HEIGHT="19"></a>&nbsp;&nbsp;
			<input type="Image" src="images/button-ff.gif" tppabs title="��һ��" name="Submit" WIDTH="73" HEIGHT="19">
		</center></div>
		</form>
	<%case "SONG"
		albumid = request("albumid")
	%>
		<form method="POST" action="song_save.asp">
		<input type="hidden" name="albumid" size="15" value="<%=albumid%>">
		<div align="center"><center>
		<table border="0" cellspacing="10" width="500" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0">
			<tr>
				<td width="100%" height="20">������� �� ������Ϣ��</td>
			</tr>
			<tr align="center">
				<td width="100%">
					<table border="0" cellspacing="1" width="100%">
						<tr>
							<td width="20%" align="right">�����ļ���</td>
							<td width="80%">
								<input type="text" name="songfile" size="25" class="smallinput" maxlength="100" value>
							</td>
						</tr>
						<tr>
							<td width="20%" align="right">�������ƣ�</td>
							<td width="80%">
								<input type="text" name="songname" size="25">
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
								&nbsp;&nbsp;<input type="reset" value=" ������� " name="cmdcancel" style="background-color: rgb(255,255,255); border: 1px inset rgb(0,0,0)">
      		    		</td>
						</tr>
						<tr>
							<td width="20%" align="right">������ʣ�</td>
							<td width="80%"> <textarea rows="4" name="songlrc" cols="40"></textarea></td>
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
<%end select%>
<br><br>
</body>
</html>
<!--#include file="conn_music_close.asp"-->
