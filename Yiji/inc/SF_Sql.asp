<%
Dim Fy_Post,Fy_Get,Fy_In,Fy_Inf,Fy_Xh,Fy_db,Fy_dbstr
'自定义需要过滤的字串,用 "防" 分隔
websql="update|count|and|exec|insert|chr|mid|master|delete|truncate|declare|char|*|’"
Fy_In = replace(websql,"’","'")
'----------------------------------
Fy_Inf = split(Fy_In,"|")
'--------POST部份------------------
If Request.Form<>"" Then
For Each Fy_Post In Request.Form

For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh))<>0 Then
response.write"<script>alert('操作错误，下面是产生错误的可能原因：\n\n·在您提交的资料中含有敏感字符');history.go(-1);</script>"
response.end
End If
Next

Next
End If
'--------GET部份-------------------
If Request.QueryString<>"" Then
For Each Fy_Get In Request.QueryString

For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
response.write"<script>alert('操作错误，下面是产生错误的可能原因：\n\n·在您提交的资料中含有敏感字符');history.go(-1);</script>"
response.end
End If
Next
Next
End If
%>