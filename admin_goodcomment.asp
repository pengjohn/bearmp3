<!--#include file="conn_music_open.asp"-->
<%
action = request("action")
goodcomment_id=request("goodcomment_id")

set rs=server.createobject("adodb.recordset")
if action="New" then
	sql="select * from goodcomment"
	rs.open sql,conn,1,3
	rs.addnew
		rs("goodcommenttitle") 	= request.form("goodcommenttitle")
		rs("goodcommentcontent")= request.form("goodcommentcontent")
	rs.update
	rs.close
elseif action="Edit" then
    sql="select * from goodcomment where goodcomment_id="&goodcomment_id
    rs.open sql,conn,1,3
    if rs.eof and rs.bof then
    else
		rs("goodcommenttitle") 	= request.form("goodcommenttitle")
		rs("goodcommentcontent")= request.form("goodcommentcontent")
		rs.update
    end if
    rs.close
elseif action="Del" then
	conn.execute "Delete * From goodcomment Where goodcomment_id="& request.querystring("goodcomment_id") 
	response.redirect "admin_goodcomment.asp"
end if
%>

<html>
<head>
<title>骚熊音乐->首页推荐</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#000000" >
<br><br>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from goodcomment order by goodcomment_id desc;" 
rs.open sql,conn,1,1
if rs.eof and rs.bof then
else
%>
	<div align="center">
	<table border="1" width="500" bordercolor="#555555">
		<%Do While Not rs.EOF or err%>   
		<tr>
			<td width="70%"><%=rs("goodcommenttitle")%></td>
			<td width="15%"><a href=admin_goodcomment.asp?goodcomment_id=<%=rs("goodcomment_id")%>#添加编辑>编辑</a></td>
			<td width="15%"><a href=admin_goodcomment.asp?action=Del&goodcomment_id=<%=rs("goodcomment_id")%>>删除</a></td>
		</tr>
		<%rs.movenext
		loop
		rs.close
end if%>
</table>
	</div>
	<%
	if goodcomment_id<>"" then
		set rs=server.createobject("adodb.recordset")
		sql="select * from goodcomment where goodcomment_id="&goodcomment_id
		rs.open sql,conn,1,3
		if rs.eof and rs.bof then
		else
			goodcommenttitle= rs("goodcommenttitle")
			goodcommentcontent = rs("goodcommentcontent")
    	end if
		rs.close
	end if
	%>
	<form action="admin_goodcomment.asp" method="POST">
		<%if goodcomment_id<>"" then%>
			<input type="hidden" name="action" value="Edit">
			<input type="hidden" name="goodcomment_id" value="<%=goodcomment_id%>">
		<%else%>
			<input type="hidden" name="action" value="New">
    	<%end if%>
		<table border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#FFFFFF" width="50%">
			<tr>
				<td><a href=admin_goodcomment.asp>新增推荐 </a></td>
			</tr>
			<tr>
				<td align="right"><a name="添加编辑">标题：</a></td>
				<td><input type="TEXT" name="goodcommenttitle" value='<%=goodcommenttitle%>' size="50" maxlength="50" style="background-color: rgb(233,233,233); border: 1px dotted rgb(0,0,0)"></td>
   		</tr>
   		<tr>
				<td align="right">简介：</td>
				<td><textarea name="goodcommentcontent" rows="10" cols="50" wrap="PHYSICAL" style="background-color: rgb(233,233,233); border: 1px dotted rgb(0,0,0)"><%=goodcommentcontent%></textarea></td>
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

