��<%=editpage%>��/��<%=editpagecount%>��
<%if editpage>1 then%>
| <a href="javascript:goItem(1);">��һ��</a>
| <a href="javascript:goItem(<%=editpage-1%>);">��һ��</a>
<%else%>
| ��һ��
| ��һ��
<%end if%>
<%if editpagecount>1 and editpage<editpagecount then%>
| <a href="javascript:goItem(<%=editpage+1%>);">��һ��</a>
| <a href="javascript:goItem(<%=editpagecount%>);">���һ��</a>
<%else%>
| ��һ��
| ���һ��
<%end if%>
ת����<input type="text" name="editpage" value="<%=editpage%>" style="width:30px;height:20px;">��

<script language="javascript">
	function goItem(editpage){
		document.frmitem.editpage.value=editpage;
		document.frmitem.submit();
	}
</script>