<!--#include file="conn_music_open.asp"-->
<!--#include file="conn_function.asp"-->
<%
   dim Admin
   dim sql 
   dim rs
   dim commentindex
   Admin=Session.Contents("IsAdmin")
   commentindex = request("commentindex")
%>

<html>
<head>
<title>ɧ������->��������</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#000000" >
<%
set rs=server.createobject("adodb.recordset")
sql="select * from comment where commentindex="&cstr(commentindex)&" order by comment_id desc;" 
rs.open sql,conn,1,1
if rs.eof and rs.bof then
   response.write "<br><p align='center'>���κ�����</p>"
else
   ShowContent
end if
rs.close
set rs=nothing  
%>

<%if Session.Contents("mp3_username")<>"" then%>
<div align="center">
<br>
<table border="1" width="500" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#CCCCCC" bordercolorlight="#000000">
	<tr>
		<td width="100%" height=22 background="images/dh-di.jpg" align="center">��������</td>
	</tr>
	<tr>
		<td width="100%" align="center">
			<table border="0" width="100%" cellspacing="0" cellpadding="0">
				<form method="POST" name="commentcontent" action="comment_add.asp?commentindex=<%=commentindex%>">
				<tr>
					<td width=100% valign=top colspan="2"><TEXTAREA name=commentcontent rows=12 cols="78" style="BORDER-BOTTOM: 1px solid;;font-size:9pt; BORDER-LEFT: 1px solid; BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid" title="���ݲ��ܳ���250���ַ���">���ڴ�����</TEXTAREA></td>
				</tr>
				<tr>
					<td width=100% align=center colspan="2"><input type="submit"  value="�ύ" name="B1" style="background-color: #CCCCCC; HEIGHT: 20px; POSITION: relative;font-size:9pt; border: 1 solid #000000"></td>
				</tr>
				</form>				
			</table>
		</td>
	</tr>
</table>
</div>
<%end if%>
<!--#include file="conn_music_close.asp"-->
</body>
</html>


<%Sub ShowContent%>
<div align="center">
<table border="1" width="700" bordercolor="#FFFFFF">
  <tr>
    <td width="10%" align="center">������</td>
    <td width="70%" align="center">��������</td>
    <td width="20%" align="center">����ʱ��</td>
  </tr>

<%Do While Not rs.EOF or err%>   
<tr>
    <td width="10%" align="center"><a href="mailto:<%=rs("commentemail")%>"><% =rs("commentname")%></td>
    <td width="70%"><%=changechr(rs("commentcontent"))%></a></td>
    <td width="20%" align="center"><% =rs("commentdatetime")%>
    <%if Session.Contents("IsAdmin")=true then%>
    	<br><a href="comment_del.asp?comment_id=<%=rs("comment_id")%>&commentindex=<%=commentindex%>">ɾ��</a>
    	<br>IP:<%=rs("commentip")%>
    <%end if%>
    </td>
<%
    rs.movenext
    loop
%>
</tr>
</table>
</div>
<%End Sub%>
