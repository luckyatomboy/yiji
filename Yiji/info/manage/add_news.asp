<%@language="VBscript"%>
<%Response.Expires=0%>
<link href="../../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
<%dim ThisKey
ThisKey = "d"
%>

<!--#include file="../include/checkadmin.asp"-->
<!--#include file="../include/conn.asp"-->

<style>
	td{font-size:13px;}
</style>
<script language="javascript">
	function OpenHTMLEditor(){
		window.open('./htmleditor/editor.html','','width=768,height=500,top=10,left=10');
	}
	function OpenHTMLPreview(){
		srcstr=document.all.content.value
		srcstr=srcstr.replace("[HTML]","");
		srcstr=srcstr.replace("[/HTML]","");
		PreviewPage=window.open('','_blank','');
		PreviewPage.document.write(srcstr);
	}
</script>
<script language="javascript">
	function doSubmit(){
		var obj=document.frmadd;

		if(obj.columnid.value==""){
			alert("��ѡ��Ŀ¼��");
			obj.columnid.focus();
			return false;
		}

		if(obj.title.value==""){
			alert("���ⲻ��Ϊ�գ�");
			obj.title.focus();
			return false;
		}

		if(obj.content.value==""){
			//alert("���ݲ���Ϊ�գ�");
			//obj.newsname.focus();
			//return false;
		}

		if(obj.imgname.value!=""){
			var filename;
			var fileext;
			filename=obj.imgname.value;
			fileext=filename.substr(filename.length-3,3);
			fileext=fileext.toUpperCase();
			if ((fileext != "GIF") && (fileext != "JPG")){                  
                  alert("ͼƬ��ʽ����ȷ����ѡ��GIF��JPGͼƬ��");
                  return false;
			}
		}
	}
</script>
<table border="0" cellspacing="1" width="100%" >
  <form name="frmadd" enctype="multipart/form-data" method="post" action="add_news_ok.asp" onsubmit="javascript:return doSubmit();">
  <input type="hidden" name="action" value="addnew">
  <tr>
    <td height="23" align="center"background="../../images/r_1.gif"><img src="images/spacer.gif" width="10" height="1"><font style="font-size:13px;" ><a href="list_news.asp"><b>ά������</b></a></font>
  <img src="images/spacer.gif" width="10" height="1"><font style="font-size:13px;" ><a href="add_news.asp"><b>��������</b></a></font>
  <img src="images/spacer.gif" width="10" height="1"><font style="font-size:13px;" ><a href="add_news_catalog.asp"><b>�������</b></a></font>
  <img src="images/spacer.gif" width="10" height="1"><font style="font-size:13px;" ><a href="list_review.asp"><b>���۹���</b></a></font></td>
  </tr>
  <tr>
    <td height="23" bgcolor="#FFFFFF">
	<br>
	<table align="center" border="0" width="96%">
	  <tr>
		<td height="26">Ŀ¼</td>
		<td>
		<select name="columnid">
		<option value=""></option>
		<%
		set rs_root=conn.execute("select columnid,columnname from tbl_news_column where columnparentid=0 order by columnid desc")
		do while not rs_root.eof
			response.write("<option value="&rs_root("columnid")&">��"&rs_root("columnname")&"</option>")
			set rs_child=conn.execute("select columnid,columnname from tbl_news_column where columnparentid="&rs_root("columnid")&" order by columnid desc")
			do while not rs_child.eof
				response.write("<option value="&rs_child("columnid")&"> ����"&rs_child("columnname")&"</option>")
				rs_child.movenext
			loop
			rs_root.movenext
		loop
		set rs_root=nothing
		%>
		</select>
		</td>
	  </tr>
	  <tr>
		<td height="26">����</td>
		<td><input type="text" name="title" style="width:480px;"></td>
	  </tr>
	  <tr>
		<td colspan="2" height="8"></td>
	  </tr>
	  <tr>
		<td height="26">ͼƬ</td>
		<td><input type="file" name="imgname" style="width:480px;"></td>
	  </tr>
	  <tr style="display:none;">
		<td height="26">ͼƬλ��</td>
		<td>		
		<select name="imgposition">
		<option value="0" selected>Ĭ��λ��</option>
		<option value="1">��</option>
		<option value="2">��</option>
		<option value="3">��</option>
		</select>
		</td>
	  </tr>
	  <tr>
		<td height="26">ͼƬ����</td>
		<td><input type="text" name="imgurl" style="width:410px;"></td>
	  </tr>
	  <tr>
		<td colspan="2" height="8"></td>
	  </tr>
	  <tr>
		<td height="26">���</td>
		<td>
		<input type="checkbox" name="index" value="1">��ҳ����
		<input type="checkbox" name="important" value="1">�ص�����
		</td>
	  </tr>
	  <tr>
		<td height="26" valign="top">����</td>
		<td>
		<INPUT type="hidden" name="content"> <iframe ID="content" src="../../rededit/ewebeditor.htm?id=content&style=coolblue" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe> 
		</td>
	  </tr>
	</table>
	<br>
	</td>
  </tr>
  <tr>
    <td height="23" colspan="2" align="center">
	<input type="submit" name="btnSubmit" value="ȷ��" style="width:60px;">
	<input type="reset" name="btnReset" value="ȡ��" style="width:60px;">
	</td>
  </tr>
  </form>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../../images/r_1.gif" alt="" /></td>
<td width="100%" background="../../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>��</td>
	  <td align="right">��</td>
    
    </tr>
  </table>
  </table>
