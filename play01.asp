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
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>骚熊在线播放-Copyright (C) 2002 技术中心在线音乐骚熊网, All Rights Reserved.</title>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
</head>

<body topmargin="0" leftmargin="0">
<object id="MediaPlayer1" height="145" width="388" classid="clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95">
  <param name="AudioStream" value="-1">
  <param name="AutoSize" value="0">
  <param name="AutoStart" value="-1">
  <param name="AnimationAtStart" value="-1">
  <param name="AllowScan" value="-1">
  <param name="AllowChangeDisplaySize" value="-1">
  <param name="AutoRewind" value="0">
  <param name="Balance" value="0">
  <param name="BufferingTime" value="5">
  <param name="ClickToPlay" value="-1">
  <param name="CursorType" value="0">
  <param name="CurrentPosition" value="0.0">
  <param name="CurrentMarker" value="0">
  <param name="DisplayBackColor" value="0">
  <param name="DisplayForeColor" value="16777215">
  <param name="DisplayMode" value="0">
  <param name="DisplaySize" value="0">
  <param name="Enabled" value="-1">
  <param name="EnableContextMenu" value="0">
  <param name="EnablePositionControls" value="-1">
  <param name="EnableFullScreenControls" value="0">
  <param name="EnableTracker" value="-1">
  <param name="Filename" value="<%=songfile%>">
  <param name="InvokeURLs" value="-1">
  <param name="Language" value="-1">
  <param name="Mute" value="0">
  <param name="PlayCount" value="1">
  <param name="PreviewMode" value="0">
  <param name="Rate" value="1">
  <param name="SelectionStart" value="-1">
  <param name="SelectionEnd" value="-1">
  <param name="SendOpenStateChangeEvents" value="-1">
  <param name="SendWarningEvents" value="-1">
  <param name="SendErrorEvents" value="-1">
  <param name="SendKeyboardEvents" value="0">
  <param name="SendMouseClickEvents" value="0">
  <param name="SendMouseMoveEvents" value="0">
  <param name="SendPlayStateChangeEvents" value="-1">
  <param name="ShowCaptioning" value="0">
  <param name="ShowControls" value="-1">
  <param name="ShowAudioControls" value="-1">
  <param name="ShowDisplay" value="-1">
  <param name="ShowGotoBar" value="0">
  <param name="ShowPositionControls" value="-1">
  <param name="ShowStatusBar" value="-1">
  <param name="ShowTracker" value="-1">
  <param name="TransparentAtStart" value="0">
  <param name="VideoBorderWidth" value="0">
  <param name="VideoBorderColor" value="0">
  <param name="VideoBorder3D" value="0">
  <param name="Volume" value="0">
  <param name="WindowlessVideo" value="0">
</object>
</body>
</html>
<!--#include file="conn_music_close.asp"-->