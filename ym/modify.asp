<!--#include file="conn.asp"-->
<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>修改域名</title>
</head>
<body>
<%
exec="select * from yumingtab where id="&request.querystring("id")
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1
%>
<form name="form1" method="post" action="modifysave.asp">
<div width="628" align="center">
    <lable>【域名信息修改】</lable>
    <p>域名<input type="text" name="yuming" value="<%=rs("yuming")%>"></p>
    <p>状态<input type="text" name="ymzhuangtai" value="<%=rs("ymzhuangtai")%>"></p>
    <input type="submit" name="Submit" value="确定修改">
    <input type="hidden" name="id" value="<%=request.querystring("id")%>">
</div>
</form>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</body>
</html>