<%
if Session.Contents("mp3_UserName")="" then
	response.write("SORRY！您没有<a href=index.htm>登录</a>。")
else
	if request.form("gstbooktitle")="" then%>
		<script language=vbscript>
			MsgBox "请写上帖子的主题！"
			location.href = "javascript:history.back()"
		</script>
	<%else%>
		<!--#include file="conn_music_open.asp"-->
		<%
		dim action
		dim gstbooktitle
		dim gstbookcontent
		dim gstbookTopID
		dim gstbooklayer
		dim gstbookorder
		dim gstbooklasttime
		dim rs,sql

		action = request("action")
		gstbooktitle = request("mode")&request("gstbooktitle")
		gstbookcontent = request("gstbookcontent")
		gstbookTopID = request("gstbookTopID")
		gstbooklayer = request("gstbooklayer")
		gstbookorder = request("gstbookorder")
		gstbooklasttime = now()

'		if request.form("gstbookcontent")="" then
'			if len(trim(request.form("gstbooktitle")))>30 then
'				gstbooktitle = left(trim(request.form("gstbooktitle")),30) & "…[无内容]"
'			else
'				gstbooktitle = trim(request.form("gstbooktitle")) & " [无内容]"
'			end if      
'		else
'			if len(trim(request.form("gstbooktitle")))>30 then
'				gstbooktitle = left(trim(request.form("gstbooktitle")),30) & "……"
'			else
'				gstbooktitle = trim(request.form("gstbooktitle")) 
'			end if      
'		end if

		set rs=server.createobject("adodb.recordset")
		select case action
			case "new"	'添加新帖子
				sql="select * from gstbook order by gstbook_id desc" 
				rs.open sql,conn,1,3
				if rs.eof and rs.bof then
					gstbookTopID = 1
				else
					gstbookTopID = rs("gstbook_id")+1
				end if
				rs.close

				sql="select * from gstbook where (gstbook_id is null)" 
				rs.open sql,conn,1,3
				rs.addnew
					rs("gstbookdatetime")  = now()
					rs("gstbooklasttime")  = now()
					rs("gstbookname") = Session.Contents("mp3_username")
					rs("gstbooktitle") = gstbooktitle
					rs("gstbookcontent") = gstbookcontent
					rs("gstbookTopID") = gstbookTopID
					rs("gstbookorder") = 1
					rs("gstbooklayer") = 1
				rs.update
				rs.close
				set rs=nothing
				conn.close
				set conn=nothing
				response.redirect("gstbook.asp")
			case "follow"	'添加跟随帖子
				sql="select * from gstbook where gstbookTopID = "&gstbookTopID
				rs.open sql,conn,1,3
				do while not rs.eof
					rs("gstbooklasttime") = gstbooklasttime
					rs.update
				rs.movenext
				loop
				rs.close

				sql="select * from gstbook where gstbookorder >= "&gstbookorder
				rs.open sql,conn,1,3
				do while not rs.eof
					rs("gstbookorder") = rs("gstbookorder")+1
					rs.update
				rs.movenext
				loop
				rs.close

				sql="select * from gstbook where (gstbook_id is null)" 
				rs.open sql,conn,1,3
				rs.addnew
					rs("gstbookdatetime")  = gstbooklasttime
					rs("gstbooklasttime")  = gstbooklasttime
					rs("gstbookname") = Session.Contents("mp3_username")
					rs("gstbooktitle") = trim(request.form("gstbooktitle"))
					rs("gstbookcontent") = gstbookcontent
					rs("gstbookTopID") = gstbookTopID
					rs("gstbookorder") = gstbookorder
					rs("gstbooklayer") = gstbooklayer
				rs.update
				rs.close
				set rs=nothing

				conn.close
				set conn=nothing
				response.redirect("gstbook.asp")
			case "modiplan"	'修改帖子
				conn.execute("Update bbs Set Topic='"& left(trim(request.form("topic")),30)&"',[text]='"&request.form("text")&"',Face="& request.form("Face") &",Midi="&request.form("Mid")&" WHERE Owner='" & owner & "' And ID=" & id)

				conn.close
				set conn=nothing
				response.redirect("gstbook_disp.asp?owner=" & owner &"&ID="&id)
		end select
	end if
end if

%>