<%
Dim Fy_Post,Fy_Get,Fy_In,Fy_Inf,Fy_Xh,Fy_db,Fy_dbstr
'�Զ�����Ҫ���˵��ִ�,�� "��" �ָ�
websql="update|count|and|exec|insert|chr|mid|master|delete|truncate|declare|char|*|��"
Fy_In = replace(websql,"��","'")
'----------------------------------
Fy_Inf = split(Fy_In,"|")
'--------POST����------------------
If Request.Form<>"" Then
For Each Fy_Post In Request.Form

For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh))<>0 Then
response.write"<script>alert('�������������ǲ�������Ŀ���ԭ��\n\n�������ύ�������к��������ַ�');history.go(-1);</script>"
response.end
End If
Next

Next
End If
'--------GET����-------------------
If Request.QueryString<>"" Then
For Each Fy_Get In Request.QueryString

For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
response.write"<script>alert('�������������ǲ�������Ŀ���ԭ��\n\n�������ύ�������к��������ַ�');history.go(-1);</script>"
response.end
End If
Next
Next
End If
%>