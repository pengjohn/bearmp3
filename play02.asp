<!--#include file="conn_music_open.asp"-->
<%
song_id = request("song_id")
set rs=server.CreateObject("ADODB.RecordSet")
sql = "select * from song where song_id="&song_id
rs.open sql,conn,1,3
	songfile = rs("songfile")
	rs("songhit") = rs("songhit")+1
	rs.update
rs.close
set rs=nothing
%>
<HTML>
<HEAD>
<TITLE>ÔÚÏß¸è´ÊVer.0.1Beta°æ </TITLE>
<style type="text/css">
<!--
	Body,td {
		FONT-FAMILY: "Verdana", "Arial", "Helvetica", "sans-serif";
		COLOR: #00FF00;
		FONT-SIZE: 12px;}
	A:link {
		COLOR: #00FF00;
		TEXT-DECORATION: none;} 
	A:visited {
		COLOR: #00FF00;
		TEXT-DECORATION: none;}
	A:active {
		COLOR: #00FF00;
		TEXT-DECORATION: none;} 
	A:hover {
		COLOR: #00FF00;
		TEXT-DECORATION: underline;} 
	Input {
		FONT-SIZE: 12px;
		COLOR: #00FF00;
		BACKGROUND-COLOR: #255A95;
		HEIGHT: 15px;
		BORDER-LEFT: 0;
		BORDER-RIGHT: 0;
		BORDER-TOP: 0;
		BORDER-BOTTOM: 1px solid #00FF00;
		FONT-FAMILY: "Verdana", "Arial", "Helvetica", "sans-serif";}
-->
</style>
</HEAD>

<BODY topmargin="0" leftmargin="0" bgcolor="#255A95">

<center>
<iframe src="weblrc/Lyrics.html" width="0" height="0" scrollbar="no" scrolling="no" name="WebLrc" border="0" frameborder="0"></iframe>
<br>
<iframe src="weblrc/WmpPlayer.asp?filename='<%=songfile%>'" width="375" height="100" scrollbar="no" scrolling="no" name="WebPlayer" border="0" frameborder="0"></iframe>
<br><br>
<br><br>
</center>

</BODY>
</HTML>
<!--#include file="conn_music_close.asp"-->