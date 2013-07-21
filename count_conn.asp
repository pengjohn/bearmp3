<%
    dim conn_count
    dim DBpath_count
    set conn_count=server.createobject("adodb.connection")
    DBPath_count = Server.MapPath("pengjohncount.asp")		'数据库文件
    conn_count.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath_count
%>