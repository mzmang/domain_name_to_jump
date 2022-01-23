<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>访问列表</title>
</head>
<body>
<div width="628" align="center">
<lable>【访问信息列表】</lable>
<%
exec="select * from iptablist"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1

%>
<table width="628" height="24" border="1" align="center" cellpadding="1" cellspacing="0">
<tr>
<td>信息ID</td>
<td>IP地址</td>
<td>访问时间</td>
</tr>
<%
if rs.eof and rs.bof then
response.write("暂时无访问信息")
else
do while not rs.eof
%>


<tr>
<td width="66" height="22" ><%=rs("id")%></td>
<td width="66" ><%=rs("IPText")%></td>
<td width="120" ><%=rs("IPTime")%></td>
</tr>
<%
rs.movenext
loop
end if
%>

<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<div width="628" height="24" align="center" style="padding-top:20px;">
<input name="add" type="button" value="添加域名" onClick="location.href='add.asp'">
</div>
</div>
</body>
</html>