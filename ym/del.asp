<!--#include file="conn.asp"-->
<%
exec="delete * from yumingtab where id="&request.querystring("id")
conn.execute exec
conn.close
set conn=nothing
response.redirect "index.asp"
%>