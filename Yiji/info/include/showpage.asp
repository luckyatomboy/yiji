��<%=intPage%>ҳ/��<%=intPagecount%>ҳ
| ��<%=intRecordcount%>��
<%if intPage>1 then%>
| <a href="javascript:goPage(1);">��ҳ</a>
| <a href="javascript:goPage(<%=intPage-1%>);">��ҳ</a>
<%else%>
| ��ҳ
| ��ҳ
<%end if%>
<%if intPagecount>1 and intPage<intPagecount then%>
| <a href="javascript:goPage(<%=intPage+1%>);">��ҳ</a>
| <a href="javascript:goPage(<%=intPagecount%>);">βҳ</a>
<%else%>
| ��ҳ
| βҳ
<%end if%>
ת����<input type="text" name="intPage" value="<%=intPage%>" style="width:30px;height:20px;">ҳ
<input type="submit" name="btnSubmit" value="GoTo" style="width:40px;">
<script language="javascript">
	function goPage(intPage){
		document.frmpage.intPage.value=intPage;
		document.frmpage.submit();
	}
</script>