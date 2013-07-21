<!--#include file="conn_music_open.asp"-->
<%
action = request("action")
software_id=request("software_id")

set rs=server.createobject("adodb.recordset")
if action="New" then
	sql="select * from software"
	rs.open sql,conn,1,3
	rs.addnew
		rs("software_title") 	= request.form("software_title")
		rs("software_dl")= request.form("software_dl")
		rs("software_memo") 	= request.form("software_memo")
	rs.update
	rs.close
elseif action="Edit" then
    sql="select * from software where software_id="&software_id
    rs.open sql,conn,1,3
    if rs.eof and rs.bof then
    else
		rs("software_title") 	= request.form("software_title")
		rs("software_dl")= request.form("software_dl")
		rs("software_memo") 	= request.form("software_memo")
		rs.update
    end if
    rs.close
elseif action="Del" then
	conn.execute "Delete * From software Where software_id="& request.querystring("software_id") 
	response.redirect "admin_software.asp"
end if
%>

<html>
<head>
<title>骚熊音乐->首页链接</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#000000" >
<br><br>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from software order by software_id desc;" 
rs.open sql,conn,1,1
if rs.eof and rs.bof then
else
%>
	<div align="center">
	<table border="1" width="500" bordercolor="#555555">
		<%Do While Not rs.EOF or err%>   
		<tr>
			<td width="20%"><%=rs("software_title")%></td>
			<td width="40%"><%=rs("software_dl")%></td>
			<td width="20%"><%=rs("software_memo")%></td>
			<td width="10%"><a href=admin_software.asp?software_id=<%=rs("software_id")%>#添加编辑>编辑</a></td>
			<td width="10%"><a href=admin_software.asp?action=Del&software_id=<%=rs("software_id")%>>删除</a></td>
		</tr>
		<%rs.movenext
		loop
		rs.close
end if%>
	</table>
	</div>
	<%
	if software_id<>"" then
		set rs=server.createobject("adodb.recordset")
		sql="select * from software where software_id="&software_id
		rs.open sql,conn,1,3
			if rs.eof and rs.bof then
			else
				software_title= rs("software_title")
				software_dl = rs("software_dl")
				software_memo = rs("software_memo")
    	end if
    rs.close
	end if
	%>
	<form action="admin_software.asp" method="POST">
		<%if software_id<>"" then%>
			<input type="hidden" name="action" value="Edit">
			<input type="hidden" name="software_id" value="<%=software_id%>">
		<%else%>
			<input type="hidden" name="action" value="New">
    	<%end if%>
		<table border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#FFFFFF" width=500>
			<tr>
				<td align="right"><a name="添加编辑">软件名称：</a></td>
				<td><input type="TEXT" name="software_title" value='<%=software_title%>' size="50" maxlength="50" style="background-color: rgb(233,233,233); border: 1px dotted rgb(0,0,0)"></td>
   		</tr>
   		<tr>
				<td align="right">下载地址：</td>
				<td><input type="TEXT" name="software_dl" value='<%=software_dl%>' size="50" maxlength="100" style="background-color: rgb(233,233,233); border: 1px dotted rgb(0,0,0)"></td>
   		</tr>
   		<tr>
				<td align="right">简介：</td>
				<td><textarea name="software_memo" rows="8" cols="50" wrap="PHYSICAL" style="background-color: rgb(233,233,233); border: 1px dotted rgb(0,0,0)"><%=software_memo%></textarea></td>
   		</tr>
   		<tr>
      		<td align="center" colspan="2">
					<input type="SUBMIT" name="Submit" value="确定" style="background-color: rgb(233,233,233); border: 1px dotted rgb(0,0,0)">
					<input type="RESET" name="Reset" value="清除" style="background-color: rgb(233,233,233); border: 1px dotted rgb(0,0,0)">
		      </td>
    		</tr>
		</table>
	</form>
<br><br>
</body>
</html>
<%set rs=nothing%>
<!--#include file="conn_music_close.asp"-->

