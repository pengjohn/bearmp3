<!--#include file="conn_music_open.asp"-->
<%
action = request("action")
link_id=request("link_id")

set rs=server.createobject("adodb.recordset")
if action="New" then
	sql="select * from link"
	rs.open sql,conn,1,3
	rs.addnew
		rs("link_title") 	= request.form("link_title")
		rs("link_address")= request.form("link_address")
		rs("link_memo") 	= request.form("link_memo")
	rs.update
	rs.close
elseif action="Edit" then
    sql="select * from link where link_id="&link_id
    rs.open sql,conn,1,3
    if rs.eof and rs.bof then
    else
		rs("link_title") 	= request.form("link_title")
		rs("link_address")= request.form("link_address")
		rs("link_memo") 	= request.form("link_memo")
		rs.update
    end if
    rs.close
elseif action="Del" then
	conn.execute "Delete * From link Where link_id="& request.querystring("link_id") 
	response.redirect "admin_link.asp"
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
sql="select * from link order by link_id desc;" 
rs.open sql,conn,1,1
if rs.eof and rs.bof then
else
%>
	<div align="center">
	<table border="1" width="500" bordercolor="#555555">
		<%Do While Not rs.EOF or err%>   
		<tr>
			<td width="20%"><%=rs("link_title")%></td>
			<td width="40%"><%=rs("link_address")%></td>
			<td width="20%"><%=rs("link_memo")%></td>
			<td width="10%"><a href=admin_link.asp?link_id=<%=rs("link_id")%>#添加编辑>编辑</a></td>
			<td width="10%"><a href=admin_link.asp?action=Del&link_id=<%=rs("link_id")%>>删除</a></td>
		</tr>
		<%rs.movenext
		loop
		rs.close
end if%>
</table>
	</div>
	<%
	if link_id<>"" then
		set rs=server.createobject("adodb.recordset")
		sql="select * from link where link_id="&link_id
		rs.open sql,conn,1,3
			if rs.eof and rs.bof then
			else
				link_title= rs("link_title")
				link_address = rs("link_address")
				link_memo = rs("link_memo")
    	end if
    rs.close
	end if
	%>
	<form action="admin_link.asp" method="POST">
		<%if link_id<>"" then%>
			<input type="hidden" name="action" value="Edit">
			<input type="hidden" name="link_id" value="<%=link_id%>">
		<%else%>
			<input type="hidden" name="action" value="New">
    	<%end if%>
		<table border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#FFFFFF" width="50%">
			<tr>
				<td><a href=admin_link.asp>新增推荐 </a></td>
			</tr>
			<tr>
				<td align="right"><a name="添加编辑">站名：</a></td>
				<td><input type="TEXT" name="link_title" value='<%=link_title%>' size="15" maxlength="50" style="background-color: rgb(233,233,233); border: 1px dotted rgb(0,0,0)"></td>
   		</tr>
   		<tr>
				<td align="right">地址：</td>
				<td><input type="TEXT" name="link_address" value='<%=link_address%>' size="30" maxlength="100" style="background-color: rgb(233,233,233); border: 1px dotted rgb(0,0,0)"></td>
   		</tr>
   		<tr>
				<td align="right">简介：</td>
				<td><textarea name="link_memo" rows="4" cols="35" wrap="PHYSICAL" style="background-color: rgb(233,233,233); border: 1px dotted rgb(0,0,0)"><%=link_memo%></textarea></td>
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

