<%
if request.querystring("Flag")="OK" then
	response.write("<script language='javascript'>alert('������Ϣ�Ѿ��ɹ����Ƽ����������ѣ�');window.close();</script>")
	response.end
end if
if request.form("action")="news_recommend" then
	title=request.form("title")
	content=request.form("content")
	yourname=request.form("yourname")
	yourmail=request.form("yourmail")
	friendmail=request.form("friendmail")
	yourcontent=request.form("yourcontent")
	sContent="��������<b>"&yourname&"</b>�����Ƽ�������һ����Ϣ��<BR>------------------------<BR>" & content & "<BR><BR><BR>�����������ѵ����ԣ�<BR>------------------------<BR>" & yourcontent & "<BR><BR><BR><a target='_blank' href='http://www.semmw.com'>�Ϻ������</a><BR>------------------------<BR>��վ���裬��վ�ƹ㣬�������䣬ϵͳ��������ҵӦ�����..."
end if
%>
<form name="frmmail" method="post" action="sendmail.asp">
<input type="hidden" name="sFrom" value="<%=yourmail%>">
<input type="hidden" name="sTo" value="<%=friendmail%>">
<input type="hidden" name="sSubject" value="<%=title%>">
<textarea style="display:none;" name="sContent"><%=sContent%></textarea>
</form>
<script language="javascript">
	document.frmmail.submit();
</script>