<!--#include file="conn.asp"-->
<%
exec="select * from yumingtab where id="&request.form("id")
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,3
%>
<%
rs("yuming")=request.form("yuming")
rs("ymzhuangtai")=request.form("ymzhuangtai")
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
response.redirect"index.asp"
%>