<%@language="VBscript"%>
<%Response.Expires=0%>
<!--#include file="../include/checkadmin.asp"--><!--#include file="../include/conn.asp"-->
<!--#include file="../include/tools.asp"-->
<%
id=request.querystring("id")
if id="" or isnull(id) or not isnumeric(id) then
	response.write("<p align=center>��Ϣ��Ŵ���</p>")
	response.end
end if
set rs=conn.execute("select * from tbl_news where id="&id)
if rs.eof then
	response.write("<p align=center>��Ϣ�����ڻ�����Ϣ�Ѿ���ɾ����</p>")
	response.end
end if
%>
<script language="javascript">
	function doZoom(size){
		document.getElementById('zoom').style.fontSize=size+'px';
	}
	function doRecommend(id,title){//2.0������ʾ 
		mewin=window.open("news_recommend.asp?id="+id+"&title="+title,"recommend","menubar=no,toolsbar=no,scrollbars=0,top=100,left=100,width=544,height=320,location=no");
		mewin.focus();
	}
	function doReview(id){//2.0������ʾ 
		mewin=window.open("news_review.asp?id="+id,"review","menubar=no,toolsbar=no,scrollbars=0,top=100,left=100,width=544,height=420,location=no");
		mewin.focus();
	}
</script>
<style>
	td{font-size:13px;}
</style>
<title><%=rs("title")%></title>
<table align="center" border="0" width="630">
  <tr>
	<td height="26" align="center"><b style="font-size:15px;"><%=rs("title")%><b></td>
  </tr>
  <tr>
	<td height="1" bgcolor="#D6DFF7"></td>
  </tr>
  <tr>
	<td height="26" align="center">
	<%if rs("imgname")<>"" then%>
	<img src="../newsimgages/<%=rs("imgname")%>">
	<%end if%>
	</td>
  </tr>
  <tr>
	<td height="26" align="center"><%=rs("regdate")%></td>
  </tr>
  <tr>
	<td id="zoom" height="26"><%=txt2html(rs("content"))%></td>
  </tr>
  <tr>
	<td height="26" align="right">��<a href="javascript:doReview(<%=id%>);">����</a>����<a href="javascript:doZoom(16);">��</a> <a href="javascript:doZoom(14);">��</a> <a href="javascript:doZoom(12);">С</a>����<a href="javascript:window.print();">��ӡ</a>����<a href="javascript:window.close();">�ر�</a>��</td>
  </tr>
</table>
<%
conn.execute "update tbl_news set hits=hits+1 where id="&id
%>