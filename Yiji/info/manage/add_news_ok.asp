<%@language="VBscript"%>
<%Response.Expires=0%>
<!--#include file="../include/checkadmin.asp"-->
<!--#include file="../include/upload.asp"-->
<!--#include file="../include/conn.asp"-->
<%
dim upload,file,sfilename,bfilename,formName,sPath,bPath,iCount
set upload=new upload_5xSoft ''�����ϴ�����

imgpath="../newsimgages/"	'ͼƬ·��
ver="cn"

action = upload.form("action") '�༭����
if action = "addnew" then
	columnid = upload.form("columnid")
	title = replace(upload.form("title"),"'","''")
	content = upload.form("content")
	index = upload.form("index")
	if index<>"1" then index="0"
	important = upload.form("important")
	if important<>"1" then important="0"

	conn.execute "insert into tbl_news(columnid,title,index,important,content) values("&columnid&",'"&title&"','"&index&"','"&important&"','"&content&"')"
	set rs=conn.execute("select top 1 id from tbl_news order by id desc")
	if not rs.eof then
		id=rs("id")
	end if
	set rs=nothing
	'�ϴ�ͼƬ
	set file=upload.file("imgname")
	if file.FileSize>0 then
	dim fileExt
fileExt=lcase(right(file.FileName,4))
     if fileEXT<>".gif" and fileEXT<>".jpg" and fileEXT<>".swf" and fileEXT<>".bmp" then
  response.write "<font size=2>�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
 response.end
 end if 


		imgname=ver & "_" & id & right(file.FileName,4)
		file.SaveAs Server.mappath(imgpath&imgname)
		imgposition = upload.form("imgposition")
		imgurl = upload.form("imgurl")
		conn.execute "update tbl_news set imgname='"&imgname&"',imgposition='"&imgposition&"',imgurl='"&imgurl&"' where id="&id
	end if
	set file=nothing
end if
%>

<style>
	td{font-size:13px;}
</style>
<table border="0" cellspacing="1" width="100%" bgcolor="#D6DFF7">
  <tr>
    <td height="23"><img src="images/spacer.gif" width="10" height="1"><font style="font-size:13px;" color="#215DC6"><b>��������</b></font></td>
  </tr>
  <tr>
    <td height="23" bgcolor="#FFFFFF">
	<br><br>
	<p align="center">���ӳɹ����Ƿ�����������ţ�</p>
	<p align="center">
	<input type="button" name="btnOk" value="��" style="width:60px;" onclick="javascript:location.href='add_news.asp';">
	<input type="button" name="btnCancel" value="��" style="width:60px;" onclick="javascript:location.href='list_news.asp';">
	</p>
	<br>
	</td>
  </tr>
  <tr>
    <td height="23" colspan="2" align="center"></td>
  </tr>
</table>
