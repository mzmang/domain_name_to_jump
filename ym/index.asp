<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>域名列表</title>
</head>
<body>
<div width="628" align="center">
<lable>【域名信息列表】</lable>
<%
exec="select * from yumingtab"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1

%>
<table width="628" height="24" border="1" align="center" cellpadding="1" cellspacing="0">
<tr>
<td>域名ID</td>
<td>域名</td>
<td>状态</td>
<td>最新使用时间</td>
<td>操作</td>
</tr>
<%
if rs.eof and rs.bof then
response.write("暂时无可用域名")
else
do while not rs.eof
%>

<tr>
<td width="66" height="22" ><%=rs("id")%></td>
<td width="66" ><%=rs("yuming")%></td>
<td width="66" ><%=rs("ymzhuangtai")%></td>
<td width="120" ><%=rs("time")%></td>
<td width="60" ><a href="modify.asp?id=<%=rs("id")%>" target="_self">编辑</a>||<a href="del.asp?id=<%=rs("id")%>">删除</a></td>
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
</div>
</body>
</html>