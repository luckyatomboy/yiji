<%@language="VBscript"%>
<%Response.Expires=0%>
<!--#include file="../include/checkadmin.asp"--><!--#include file="../include/conn.asp"-->
<%
if request.form("action")="news_review" then
	content = request.form("content")
	yourname = request.form("yourname")
	yourmail = request.form("yourmail")
	newsid = request.form("id")
	conn.execute "insert into tbl_news_review(newsid,content,yourname,yourmail) values('"&newsid&"','"&content&"','"&yourname&"','"&yourmail&"')"
end if
%>
<script language="javascript">
	alert("�������ӳɹ���");
	self.close();
</script>