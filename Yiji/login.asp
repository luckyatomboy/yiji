<%@ Page Language="Vb"%>
<%@ Import Namespace="System.IO"%>
<HTML>
  <HEAD>
  </HEAD>
  <BODY>
<%
	dim username,password
	username=Request.Form("Text1")
	password=Request.Form("Password1")
	
	Response.Write("����û���Ϊ��" & username &"<br>")
	Response.Write("�����������Ϊ��" & password &"<br>")
%>
  </BODY>
</HTML>