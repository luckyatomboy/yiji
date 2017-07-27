
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
<title><%=dianming%> - 打印订货合同</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style2.css" rel="stylesheet" type="text/css">
<script language=javascript>
function preview() { 
bdhtml=window.document.body.innerHTML; 
sprnstr="<!--startprint-->"; 
eprnstr="<!--endprint-->"; 
prnhtml=bdhtml.substr(bdhtml.indexOf(sprnstr)+17); 
prnhtml=prnhtml.substring(0,prnhtml.indexOf(eprnstr)); 
window.document.body.innerHTML=prnhtml; 
window.print(); 
window.document.body.innerHTML=bdhtml; 
         }
</script>
</HEAD>

<BODY>
<%
  set rs_contract=conn.execute("select * from salescontract where contractnum="&request("contractnum"))
  sql="select * from owncompany where company="&rs_contract("owncompany")
  set rs_owncompany=conn.execute(sql)  
%>


<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:20pt;">
  <tr> 
    <td height="21" align="center"><img src="../images/print.jpg" align="absmiddle" style="cursor:hand;" onClick="preview();window.close()"></td>
  </tr>
</table>
<!--startprint-->

<!--logo-->
<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:20pt;">
  <tr> 
    <td width="50%" height="100" align="left"><img src="../images/qz.jpg" align="absmiddle"></td>
    <td width="50%">
      <table border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:20pt;">
        <tr>
          <td width="10%" align="left">地址：</td>
          <td width="90%" align="left"><%=rs_owncompany("fullname")%></td>
        </tr>
        <tr>
          <td width="10%" align="left">电话：</td>
          <td width="90%" align="left"><%=rs_owncompany("tel")%></td>
        </tr>    
        <tr>
          <td width="10%" align="left">传真：</td>
          <td width="90%" align="left"><%=rs_owncompany("fax")%></td>
        </tr>
        <tr>
          <td width="10%" align="left">邮箱：</td>
          <td width="90%" align="left"><%=rs_owncompany("email")%></td>
        </tr>        
      </table>
    </td>
  </tr>
</table>

<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:20pt;">
  <tr> 
    <td width="100%" align="center">
      <h2>
        <%if rs_contract("category")="A" then %>
          订货合同
        <%else%>
          期货订货合同
        <%end if%>
      </h2>
    </td>
  </tr>
</table>
<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:20pt;">
  <tr> 
    <th width="15%" height="30" align="left">甲方（卖方）：</th>
	  <td width="35%" height="30" align="left">
          <%
          sql="select * from owncompany where company="&rs_contract("owncompany")
          set rs_owncompany=conn.execute(sql)
          %>
          <%=rs_owncompany("fullname")%>
    </td>	  
	  <th width="35%" height="30" align="right">合同编号：</th>	
	  <td width="15%" height="30" align="left"><%=rs_contract("contractnum")%></td>
	</tr>		
  <tr> 
    <th width="15%" height="30" align="left">乙方（买方）：</th>
    <td width="35%" height="30" align="left">
          <%
          sql="select * from customer where customername="&rs_contract("customer")
          set rs_customer=conn.execute(sql)
          %>
          <%=rs_customer("fullname")%>
    </td>   
    <th width="35%" height="30" align="right">我司传真：</th>  
    <td width="15%" height="30" align="left"><%=rs_owncompany("fax")%></td>
  </tr> 
</table>

<p style="margin-left:20pt;">
  现甲乙双方本着友好协商、互惠互利的原则，同意按照以下条款签定本合同：
</p>

<table width="94%" border="1" cellspacing="0" cellpadding="2" align="center" bordercolor="#000000" style="margin-top:18pt;margin-left:20pt;">
  <tr>
  <th>我司编号</th>
  <th>国别</th>
	<th>厂号</th>
	<th>品名</th>
  <th>规格</th>  
  <th>包装</th>  
  <%if rs_contract("category")="A" then %>
    <th>件数</th>
  <%end if%>  
	<th>重量（吨）</th>
	<th>单价（元/吨）</th>
  <%if rs_contract("category")="A" then %>
    <th>存放冷库</th>
    <th>交货地</th>
  <%else%>
    <th>预计到港期</th>
    <th>交货港</th>      
  <%end if%>  
  </th>	
  </tr>

  <tr>
  	<td align="center"><%=rs_contract("refshipment")%></td>
  	<td align="center"><%=rs_contract("country")%></td>
  	<td align="center"><%=rs_contract("plant")%></td>
  	<td align="center"><%=rs_contract("material")%></td>
  	<td align="center"><%=rs_contract("spec")%></td>
  	<td align="center"><%=rs_contract("package")%></td>
    <%if rs_contract("category")="A" then %>
      <td align="center"><%=rs_contract("quantity")%></td>
    <%end if%>  
  	<td align="center"><%=rs_contract("weight")%></td>
  	<td align="center"><%=rs_contract("price")%></td>
    <%if rs_contract("category")="A" then %>
      <td align="center"><%=rs_contract("coldstorage")%></td>
      <td align="center"><%=rs_contract("deliveryloc")%></td>
    <%else%>
      <td align="center"><%=rs_contract("boarddate")%></td>
      <td align="center"><%=rs_contract("deliveryport")%></td>     
    <%end if%>  
  </tr>				
  <tr>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>   
    <td align="center">总价(元):</td>
    <td align="center"><%=rs_contract("totalamount")%></td>
    <td></td>  
    <td></td>  
  </tr>     
</table>

  <br/>
  <p>
    备注：此柜的柜号为： <%=rs_contract("case")%>   ，可根据柜号自行查询船期。
  </p>
  <br/>
  <p>
    付款账号：
    <br/>
      账户名称： <%=rs_owncompany("bankaccountname")%>
      账号： <%=rs_owncompany("bankaccount")%>
      开户行： <%=rs_owncompany("bankname")%>
    <% if rs_owncompany("bankaccountname2")<>"" then%>
      <br/>
      账户名称： <%=rs_owncompany("bankaccountname2")%>
      账号： <%=rs_owncompany("bankaccount2")%>
      开户行： <%=rs_owncompany("bankname2")%>
    <%end if%>
  </p>  
  <br/>	
  <p>
    付款方式：定金支付：买方需于48小时内签署此合同并回传甲方，且于2个工作日内支付定金 RMB 30000元，剩余货款 RMB <% remain=rs_contract("totalamount")-30000%><%=remain%>元于货物到港前<%if rs_contract("category")="A" then %>一<%else%>三<%end if%>天付清。以实际供应商装船时间、装船数量为准，如有误差按实际到货数量做相应调整。如买方不能如期回传本合同或支付定金，卖方有权视当时市场状态继续或取消此合同；
  </p>
  <br/>
  <p>
    交货日期：按照商检海关正常操作流程，到港后五个工作日左右可以提货。如遇采样、标签整改、卫生证信息等问题入库，冷藏费由甲方支付，交货日期顺延到海关商检放行日。
  </p>
  <br/>
  <p>
    交货方式：乙方自提。<%if rs_contract("category")="B" then %>如需原柜倒车，倒车费需乙方承担。<%end if%>
  </p>
  <br/>
  <p>
    定金说明：定金仅作为业务确认的依据，若货物到港时，客户无法赎单，需客户自行承担市场差价之损失。
  </p>
  <br/>
  <p>
    索赔条款：买方收到货品后如发现数量、物品外在性状有问题，需于到货48小时内提出；对于内在质量问题，需在到货7天内提出；索赔款项未经双方同意，不得从货款中自行扣除；乙方需配合提供详尽的索赔依据如照片、运输免责说明等等，如有必要双方可请第三方公证机关进行公证；公证费由责任方承担；我司仅受理未更改、损坏包装的产品的索赔诉求。若经合作双方协商后需要办理退换货事宜，我司只接受原包装退换货。
  </p>
  <br/>
  <p>
    争议解决：甲乙双方不得因市场和行情变动而违约，否则，另一方将保留有关追索权利；如遇战争，海损，罢工等不可抗力因素造成的无法履约，该合同自动失效，甲方退还合同定金。其他未尽事宜，双方协商解决！
    凡因本合同引起的或与本合同有关的一切争议，甲乙双方均应按合同法规定协商解决，协商不成则向上海市仲裁委员会提起仲裁，相关费用由败诉方承担。
  </p>
  <br/>
  <p>
    本合同一式两份，甲乙双方各执一份。自双方授权代表签字或盖章并收到定金之日起生效，复印件、传真件具有同样法律效力。
  </p>
  <br/>
  <p bold>
    友情提醒：期货船期会有波动，可能会顺延1-2周，我司按照规定，每周会向客户提供船期表，如未收到请及时联系以便我们重新发送；自我司发送该张船期表的两个工作日内，客户未提出异议的，视同客户认可船期表内容。
  </p>

<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:20pt;">
  <tr> 
    <th width="15%" height="30" align="left">甲方：</th>
    <td width="35%" height="30" align="left"><%=rs_owncompany("fullname")%></td>   
    <th width="35%" height="30" align="left">乙方：</th>  
    <td width="15%" height="30" align="left"><%=rs_customer("fullname")%></td>
  </tr>   
  <tr> 
    <th width="15%" height="30" align="left">甲方授权：</th>
    <td width="35%"></td>   
    <th width="35%" height="30" align="left">乙方授权：</th>  
    <td width="15%"></td>
  </tr> 
  <tr> 
    <th width="15%" height="30" align="left">日期：</th>
    <td width="35%" height="30" align="left">date()</td>   
    <th width="35%" height="30" align="left">日期：</th>  
    <td width="15%"></td>
  </tr>   
</table>

<!--endprint-->
</body>
</html>