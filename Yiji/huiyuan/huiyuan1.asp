<%
if session("redboy_username")="" then
%>
<script language="javascript">
top.location.href="../index.asp"
</script>
<%  
  response.end
end if
%>
<!-- #include file="../conn2.asp" -->
<!-- #include file="../const.asp" -->

<html>
<head>
<title><%=dianming%> - 会员管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
</HEAD>

<BODY>
<script>
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>
<%
if fla36="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">你不具备此权限，请与管理员联系！</font></center>
<%  
  response.end
end if

'取得当前页码
currentpage=request("page")
'response.write currentpage
'response.end
if currentpage<1 or currentpage="" then
  currentpage="1"
end if

'取得搜索关键字  
nowzu=request("zu") 
nowkeyword=request("keyword")
%>
<table width="100%" border="0" cellpadding="0" cellspacing="2" align="center">
<form name="form2">
  <input type="hidden" name="form" value="<%=request("form")%>">
  <input type="hidden" name="field" value="<%=request("field")%>">
  <input type="hidden" name="field2" value="<%=request("field2")%>">
  <input type="hidden" name="field3" value="<%=request("field3")%>">
  <tr> 
    <td width="5%" height="30">&nbsp;</td>
	<td width="95%" align="right">
	  会员查询：
	  <%
	  sql="select * from zu_huiyuan order by id"
	  set rs_zu=conn.execute(sql)
	  %>
	  <select name="zu" onChange="form2.submit()">
        <option value="">请选择会员组</option>
      <%
	  do while rs_zu.eof=false
	  %>
        <option value="<%=rs_zu("id")%>"<%if trim(cstr(rs_zu("id")))=nowzu then%> selected="selected"<%end if%>><%=rs_zu("zu")%></option>
      <%
	    rs_zu.movenext
	  loop
	  %>
      </select>
	  <input type="text" name="keyword" size="20" value="<%=nowkeyword%>">
	  <input type="submit" value=" 查询 " class="button">&nbsp;
	</td>
  </tr>
</form>  
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;会员管理</td>
	  <td align="right">&nbsp;</td>
    </tr>
  </table>
</td>
<td><img src="../images/r_2.gif" alt="" /></td>
</tr>
<tr>
<td></td>
<td>
<form name="form1" action="huiyuan_del.asp">
  <input type="hidden" name="form" value="<%=request("form")%>">
  <input type="hidden" name="field" value="<%=request("field")%>">
  <input type="hidden" name="field2" value="<%=request("field2")%>">
  <input type="hidden" name="field3" value="<%=request("field3")%>">
  <input type="hidden" name="page" value="<%=currentpage%>">
  <input type="hidden" name="zu" value="<%=nowzu%>">
  <input type="hidden" name="keyword" value="<%=nowkeyword%>">
  <input type="hidden" name="order1" value="<%=request("order1")%>">
  <input type="hidden" name="order2" value="<%=request("order2")%>">
  <input type="hidden" name="order3" value="<%=request("order3")%>">
  <input type="hidden" name="order4" value="<%=request("order4")%>">
  <input type="hidden" name="order5" value="<%=request("order5")%>">
  <input type="hidden" name="order6" value="<%=request("order6")%>">
  <input type="hidden" name="order7" value="<%=request("order7")%>">
  <input type="hidden" name="order8" value="<%=request("order8")%>">
  <input type="hidden" name="order9" value="<%=request("order9")%>">
  <input type="hidden" name="order10" value="<%=request("order10")%>">
  <input type="hidden" name="order11" value="<%=request("order11")%>">
  <input type="hidden" name="order12" value="<%=request("order12")%>">
  <input type="hidden" name="order13" value="<%=request("order13")%>">
  <input type="hidden" name="order14" value="<%=request("order14")%>">
  <input type="hidden" name="order15" value="<%=request("order15")%>">
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
  <tr align="center">
    <td class="category" height="30">
	  <a href="?order5=<%if request("order5")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>&field=<%=request("field")%>&field2=<%=request("field2")%>&field3=<%=request("field3")%>&page=<%=currentpage%>&zu=<%=nowzu%>&keyword=<%=nowkeyword%>" class="title">卡号<%if request("order5")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>
	</td>
    <td class="category">
	  <a href="?order1=<%if request("order1")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>&field=<%=request("field")%>&field2=<%=request("field2")%>&field3=<%=request("field3")%>&page=<%=currentpage%>&zu=<%=nowzu%>&keyword=<%=nowkeyword%>" class="title">姓名<%if request("order1")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>
	</td>
    <td class="category">
	  <a href="?order2=<%if request("order2")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>&field=<%=request("field")%>&field2=<%=request("field2")%>&field3=<%=request("field3")%>&page=<%=currentpage%>&zu=<%=nowzu%>&keyword=<%=nowkeyword%>" class="title">性别<%if request("order2")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>
	</td>
	<td class="category">
	  <a href="?order3=<%if request("order3")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>&field=<%=request("field")%>&field2=<%=request("field2")%>&field3=<%=request("field3")%>&page=<%=currentpage%>&zu=<%=nowzu%>&keyword=<%=nowkeyword%>" class="title">QQ<%if request("order3")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>
	</td>
	<td class="category">
	  <a href="?order4=<%if request("order4")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>&field=<%=request("field")%>&field2=<%=request("field2")%>&field3=<%=request("field3")%>&page=<%=currentpage%>&zu=<%=nowzu%>&keyword=<%=nowkeyword%>" class="title">入会时间<%if request("order4")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>
	</td>
	<td class="category">
	  <a href="?order6=<%if request("order6")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>&field=<%=request("field")%>&field2=<%=request("field2")%>&field3=<%=request("field3")%>&page=<%=currentpage%>&zu=<%=nowzu%>&keyword=<%=nowkeyword%>" class="title">积分<%if request("order6")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>
	</td>
	<td class="category">
	  <a href="?order7=<%if request("order7")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>&field=<%=request("field")%>&field2=<%=request("field2")%>&field3=<%=request("field3")%>&page=<%=currentpage%>&zu=<%=nowzu%>&keyword=<%=nowkeyword%>" class="title">会员组<%if request("order7")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>
	</td>
  </tr>
  <%
  set rs_huiyuan =server.createobject("ADODB.RecordSet")
  sql="select * from huiyuan where 1=1"
  if nowzu<>"" then
    sql=sql&" and id_zu="&nowzu
  end if
  if nowkeyword<>"" then
    sql=sql&" and (username like '%"&nowkeyword&"%' or card like '%"&nowkeyword&"%' or tel like '%"&nowkeyword&"%' or qq like '%"&nowkeyword&"%' or email like '%"&nowkeyword&"%')"
  end if
  
  if request("order1")<>"" then
    sql=sql&" order by username "&request("order1")
  elseif request("order2")<>"" then
    sql=sql&" order by xinbie "&request("order2")
  elseif request("order3")<>"" then
    sql=sql&" order by qq "&request("order3") 
  elseif request("order4")<>"" then
    sql=sql&" order by startdate "&request("order4") 
  elseif request("order5")<>"" then
    sql=sql&" order by card "&request("order5") 
  elseif request("order6")<>"" then
    sql=sql&" order by jifen "&request("order6")
  elseif request("order7")<>"" then
    sql=sql&" order by id_zu "&request("order7")			        
  else
    sql=sql&" order by id desc"  
  end if  
  	
  rs_huiyuan.open sql,conn,1,3
  if not rs_huiyuan.eof then
  rs_huiyuan.pagesize=maxrecord
  rs_huiyuan.absolutepage=currentpage
  for currentrec=1 to rs_huiyuan.pagesize
    if rs_huiyuan.eof then
      exit for
    end if
  %>
  <tr onMouseOver="this.className='highlight'" onMouseOut="this.className=''" <%if request("form")<>"" then%>onDblClick="window.opener.document.<%=request.querystring("form")%>.<%=request.querystring("field")%>.value='<%=rs_huiyuan("id")%>';window.opener.document.<%=request.querystring("form")%>.<%=request.querystring("field2")%>.value='<%=rs_huiyuan("username")%>';window.opener.document.<%=request.querystring("form")%>.<%=request.querystring("field3")%>.value='<%=rs_huiyuan("card")%>';window.close();"<%else%>onDblClick="javascript:var win=window.open('huiyuan_show.asp?id=<%=rs_huiyuan("id")%>','会员详细信息','width=853,height=470,top=176,left=161,toolbar=no,location=no,directories=no,status=no,menubar=no,resizable=no,scrollbars=yes'); win.focus()"<%end if%>>
	<td align="center" height="25"><%=rs_huiyuan("card")%></td>  
    <td align="center"><%=rs_huiyuan("username")%></td>
	<td align="center"><%=rs_huiyuan("xinbie")%></td>
	<td align="center"><%if rs_huiyuan("qq")<>"" then%><%=rs_huiyuan("qq")%><a href="http://wpa.qq.com/msgrd?V=1&Uin=<%=rs_huiyuan("qq")%>&Site=<%=dianming%>&Menu=yes" target="_blank"><img border="0" SRC=http://wpa.qq.com/pa?p=1:<%=rs_huiyuan("qq")%>:7 alt="<%=rs_huiyuan("qq")%>"></a><%else%>&nbsp;<%end if%></td>
	<td align="center">
	<%if rs_huiyuan("enddate")-date()<0 then%>
	  <font color="#ff0000"><b><%=rs_huiyuan("startdate")%></b></font>
	<%elseif rs_huiyuan("enddate")-date()<10 then%>
	  <font color="#009900"><b><%=rs_huiyuan("startdate")%></b></font>
	<%else%>
	  <%=rs_huiyuan("startdate")%>
	<%end if%>
	</td>
    <td align="center"><%=rs_huiyuan("jifen")%></td>
	<td align="center">
	  <%set rs_zu=conn.execute("select * from zu_huiyuan where id="&rs_huiyuan("id_zu"))%>
	  <%if rs_zu.eof then%>&nbsp;<%else%><%=rs_zu("zu")%><%end if%>
	</td>
  </tr>
  <%
    rs_huiyuan.movenext
  next
  else
  %>
  <tr align="center" onMouseOver="this.className='highlight'" onMouseOut="this.className=''">
    <td colspan="9" height="25" align="center" style="color:red"><b>没有找到记录</b></td>
  </tr>
  <%
  end if
  %>
  
  
  
  <%
  if rs_huiyuan.recordcount>0 then 
  %> 
  <tr>
    <td colspan="9" height="30" class="category">
	<table cellpadding=0 cellspacing=0 width="100%">
	<tr>
	<td align="left" style="color:#FF0000;">&nbsp;<%if request("form")<>"" then%>双击选择会员<%else%>双击每行可查看会员详细资料<%end if%></td>  
	<td align="right">
	  &nbsp;&nbsp;
      <%=rs_huiyuan.recordcount%>&nbsp;条信息&nbsp; 共&nbsp;<%=rs_huiyuan.pagecount%>&nbsp;页&nbsp;
  <%
  nowstart=currentpage-3
  if nowstart<1 then
    nowstart=1
  end if
  nowend=currentpage+3
  if nowend>rs_huiyuan.pagecount then
    nowend=rs_huiyuan.pagecount
  end if
  response.write "&nbsp;<a href='?form="&request("form")&"&field="&request("field")&"&field2="&request("field2")&"&field3="&request("field3")&"&zu="&nowzu&"&keyword="&nowkeyword&"&page=1&order1="&request("order1")&"&order2="&request("order2")&"&order3="&request("order3")&"&order4="&request("order4")&"&order5="&request("order5")&"&order6="&request("order6")&"&order7="&request("order7")&"&order8="&request("order8")&"&order9="&request("order9")&"&order10="&request("order10")&"&order11="&request("order11")&"&order12="&request("order12")&"&order13="&request("order13")&"&order14="&request("order14")&"&order15="&request("order15")&"' class='page'>最前页</a>&nbsp;"
  for ipage=nowstart to nowend
    if cstr(ipage)=cstr(currentpage) then
      response.write "&nbsp;<span style='font-weight:bold;color:#5858E6'>" & ipage &"</span>&nbsp;"
    else
      response.write "&nbsp;[&nbsp;<a href='?form="&request("form")&"&field="&request("field")&"&field2="&request("field2")&"&field3="&request("field3")&"&zu="&nowzu&"&keyword="&nowkeyword&"&page="&ipage&"&order1="&request("order1")&"&order2="&request("order2")&"&order3="&request("order3")&"&order4="&request("order4")&"&order5="&request("order5")&"&order6="&request("order6")&"&order7="&request("order7")&"&order8="&request("order8")&"&order9="&request("order9")&"&order10="&request("order10")&"&order11="&request("order11")&"&order12="&request("order12")&"&order13="&request("order13")&"&order14="&request("order14")&"&order15="&request("order15")&"' class='page'>" & ipage &"</a>&nbsp;]&nbsp;"
    end if
  next
  response.write "&nbsp;<a href='?form="&request("form")&"&field="&request("field")&"&field2="&request("field2")&"&field3="&request("field3")&"&zu="&nowzu&"&keyword="&nowkeyword&"&page="&rs_huiyuan.pagecount&"&order1="&request("order1")&"&order2="&request("order2")&"&order3="&request("order3")&"&order4="&request("order4")&"&order5="&request("order5")&"&order6="&request("order6")&"&order7="&request("order7")&"&order8="&request("order8")&"&order9="&request("order9")&"&order10="&request("order10")&"&order11="&request("order11")&"&order12="&request("order12")&"&order13="&request("order13")&"&order14="&request("order14")&"&order15="&request("order15")&"' class='page'>最后页</a>&nbsp;"
%>
  </td>
  </tr></table></td></tr>
<%end if%> 
</table>
</form>
</td>
<td></td>
</tr>
<tr>
<td><img src="../images/r_4.gif" alt="" /></td>
<td></td>
<td><img src="../images/r_3.gif" alt="" /></td>
</tr>
</table>
</body>
</html>