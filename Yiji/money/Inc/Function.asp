<%
Dim i,text,checkpage,page,MaxPerPage,CurrentPage
Sub showpage
Response.Write "<table width='98%' border='0' cellpadding='0' cellspacing='0' align='center'>"
Response.Write "<tr> "
Response.Write "<td height='25' colspan='5' align='center'>&nbsp;&nbsp; "
Response.write "<strong>��</font>" & "<font color=#FF0000>" & Cstr(rs.RecordCount) & "</font>" & "<font color='#000000'>����¼</font></strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "<strong><font color='#000000'>��</font>" & "<font color=#FF0000>" & Cstr(CurrentPage) & "</font>" & "<font color='#000000'>/" & Cstr(rs.pagecount) & "</font></strong>&nbsp;"

If currentpage > 1 Then
response.write "<strong><a href='?&page="+cstr(1)+"'><font color='#000000'>��ҳ</font></a><font color='#ffffff'> </font></strong>"
Response.write "<strong><a href='?page="+Cstr(currentpage-1)+"'><font color='#000000'>��һҳ</font></a><font color='#ffffff'> </font></strong>"
Else
Response.write "<strong><font color='#000000'>��һҳ </font></strong>"
End if

If currentpage < rs.PageCount Then
Response.write "<strong><a href='?page="+Cstr(currentPage+1)+"'><font color='#000000'>��һҳ</font></a><font color='#ffffff'> </font>"
Response.write "<a href='?page="+Cstr(rs.PageCount)+"'><font color='#000000'>βҳ</font></a></strong>&nbsp;&nbsp;"
Else
Response.write ""
Response.write "<strong><font color='#000000'>��һҳ</font></strong>&nbsp;&nbsp;"
End if
Response.write "</td>"
Response.write "</tr></table>"
End Sub

%>





