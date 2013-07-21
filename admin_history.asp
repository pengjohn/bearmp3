<!--#include file="conn_music_open.asp"-->
<%
action = request("action")
history_id=request("history_id")

set rs=server.createobject("adodb.recordset")
if action="New" then
	sql="select * from history"
	rs.open sql,conn,1,3
	rs.addnew
		rs("historycontent") 	= request.form("historycontent")
		rs("historydate")= now()
	rs.update
	rs.close
elseif action="Edit" then
    sql="select * from history where history_id="&history_id
    rs.open sql,conn,1,3
    if rs.eof and rs.bof then
    else
		rs("historycontent") 	= request.form("historycontent")
		rs.update
    end if
    rs.close
elseif action="Del" then
	conn.execute "Delete * From history Where history_id="& request.querystring("history_id") 
	response.redirect "admin_history.asp"
end if
%>

<html>
<head>
<title>骚熊音乐->首页公告</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#000000" >
<br><br>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from history order by history_id desc;" 
rs.open sql,conn,1,1
if rs.eof and rs.bof then
else
%>
	<div align="center">
	<table border="1" width="600" bordercolor="#555555">
		<%Do While Not rs.EOF or err%>   
		<tr>
			<td width="60%"><%=rs("historycontent")%></td>
			<td width="20%"><%=rs("historydate")%></td>
			<td width="10%"><a href=admin_history.asp?lhistory_id=<%=rs("history_id")%>#添加编辑>编辑</a></td>
			<td width="10%"><a href=admin_history.asp?action=Del&history_id=<%=rs("history_id")%>>删除</a></td>
		</tr>
		<%rs.movenext
		loop
		rs.close
end if%>
</table>
	</div>
	<%
	if history_id<>"" then
		set rs=server.createobject("adodb.recordset")
		sql="select * from history where history_id="&history_id
		rs.open sql,conn,1,3
			if rs.eof and rs.bof then
			else
				historycontent = rs("historycontent")
    	end if
    rs.close
	end if
	%>
	<form action="admin_history.asp" method="POST">
		<%if history_id<>"" then%>
			<input type="hidden" name="action" value="Edit">
			<input type="hidden" name="history_id" value="<%=history_id%>">
		<%else%>
			<input type="hidden" name="action" value="New">
    	<%end if%>
		<table border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#FFFFFF" width="50%">
			<tr>
				<td><a href=admin_history.asp>新增推荐 </a></td>
			</tr>
   		<tr>
				<td align="right">更新内容：</td>
				<td><textarea name="historycontent" rows="4" cols="35" wrap="PHYSICAL" style="background-color: rgb(233,233,233); border: 1px dotted rgb(0,0,0)"><%=historycontent%></textarea></td>
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

