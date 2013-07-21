<%   dim conn
   dim DBpath
   set conn=server.createobject("adodb.connection")
   DBPath = Server.MapPath("pengjohnmusic.asp")
   conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
%>
