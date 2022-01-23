<!--#include file="conn.asp"-->

<%
yuming=request.form("yuming")
exec="insert into yumingtab(yuming) values('"+yuming+"')"
conn.execute exec
conn.close
set conn=nothing
response.redirect "index.asp"
%>