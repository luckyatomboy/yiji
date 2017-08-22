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
<title><%=dianming%> - 修改船期表</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script type="text/javascript" src="../js/jquery-3.2.1.js"></script>
<script type="text/javascript" src="../js/jquery-ui.js"></script>
<link href="../style/jquery-ui.css" rel="stylesheet" type="text/css">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
</HEAD>

<BODY>
<%
if fla1="0" and fla2="0" and fla3="0" and fla4="0" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">你不具备此权限，请与管理员联系！</font></center>
<%  
  response.end
end if
%>


<%
sql="select * from shipment where shipmentnum="&request("shipment")
set rs=conn.execute(sql)
%>

<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td width="10%" align="left">&nbsp;查看船期表资料  ---  </td>
	  <td align="left"><%=rs("shipmentnum")%></td>
    </tr>
  </table>
</td>
<td><img src="../images/r_2.gif" alt="" /></td>
</tr>
<tr>
<td></td>
<td>
	  <form name="form1">	
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
<!--	  <form name="form1">	-->
      <tr>
        <td width="20%" align="right" height="30">船期表号码：</td>
        <td width="80%" class="category"><input name="shipment" id="shipment" style="border:none" value=<%=rs("shipmentnum")%> readonly></td>
      </tr>
      <tr>
        <td width="20%" align="right" height="30">一般信息：</td>
        <td width="80%" class="category">
        	类型：
					<%
					sql="select * from trantype"
					set rs_trantype=conn.execute(sql)
					%>
					<select name="trantype" readonly>
						<%do while rs_trantype.eof=false%>
						<option value="<%=rs_trantype("trantype")%>" <%if rs_trantype("trantype")=rs("trantype") then%>selected="selected"<%end if%>><%=rs_trantype("trantype")%></option>
						<%
							rs_trantype.movenext
							loop
						%>
			</select>
					&nbsp;&nbsp;&nbsp;&nbsp;状态
					<%
					sql="select * from status"
					set rs_status=conn.execute(sql)
					%>
					<select name="status" readonly>
						<%do while rs_status.eof=false%>
						<option value="<%=rs_status("status")%>" <%if rs_status("status")=rs("status") then%>selected="selected"<%end if%>><%=rs_status("status")%></option>
						<%
							rs_status.movenext
							loop
						%>
					</select>
					&nbsp;&nbsp;&nbsp;&nbsp;结算单状态
					<%
					sql="select * from invoicestatus"
					set rs_invoicestatus=conn.execute(sql)
					%>
					<select name="invoicestatus" readonly>
						<%do while rs_invoicestatus.eof=false%>
						<option value="<%=rs_invoicestatus("status")%>" <%if rs_invoicestatus("status")=rs("invoicestatus") then%>selected="selected"<%end if%>><%=rs_invoicestatus("status")%></option>
						<%
							rs_invoicestatus.movenext
							loop
						%>
					</select>					
			  </td>
			</tr>
      <tr>	  
	    	<td align="right" height="30">采购：</td>
        <td class="category">		
					<select name="buyer" readonly>
	    			<% for i = 0 to buyerCount-1 %>
	 						<option value="<%=buyer(i)%>" <%if buyer(i)=rs("buyer") then%>selected="selected"<%end if%>><%=buyer(i)%></option>
	 					<% next %>  			
					</select>						
					&nbsp;&nbsp;&nbsp;&nbsp;跟单	      
					<select name="custom" readonly>
	    			<% for i = 0 to customCount-1 %>
	 						<option value="<%=custom(i)%>" <%if custom(i)=rs("handler") then%>selected="selected"<%end if%>><%=custom(i)%></option>
	 					<% next %>  			
					</select>				
					&nbsp;&nbsp;&nbsp;&nbsp;销售	      
					<select name="sales" readonly>
	    			<% for i = 0 to salesCount-1 %>
	 						<option value="<%=sales(i)%>" <%if sales(i)=rs("sales") then%>selected="selected"<%end if%>><%=sales(i)%></option>
	 					<% next %>  			
					</select>									  		
				</td>
      </tr>			
      <tr>	  
	    	<td align="right" height="30">供应商：</td>
        <td class="category">
					<%
					sql="select * from vendor order by vendorname"
					set rs_vendor=conn.execute(sql)
					%>
					<select name="vendor" readonly>
						<%
							do while rs_vendor.eof=false
						%>
							<option value="<%=rs_vendor("vendorname")%>" <%if rs_vendor("vendorname")=rs("vendor") then%>selected="selected"<%end if%>><%=rs_vendor("vendorname")%></option>
						<%
							rs_vendor.movenext
							loop
							rs_vendor.close
						%>
					</select>
					&nbsp;<font color="#ff0000">*</font>
					&nbsp;&nbsp;&nbsp;&nbsp;国别				
					<select name="country" readonly>			
						<option selected="selected"><%=rs("country")%></option>		  			
					</select>						
					&nbsp;&nbsp;&nbsp;&nbsp;厂号						
					<select name="plant" readonly>
						  <option selected="selected"><%=rs("plant")%></option>		
					</select>					
		  		&nbsp;&nbsp;&nbsp;付款方式
					<select name="incoterm" readonly>			
						  <option selected="selected"><%=rs("incoterm")%></option>										
					</select>				  				
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">合同号：</td>
        <td class="category">
		  		<input name="contract" value=<%=rs("contract")%> style="width:150px;border:none" readonly maxlength="20">
		  		&nbsp;<font color="#ff0000">*</font>
		  		&nbsp;&nbsp;&nbsp;&nbsp;品类
					<%
					sql="select * from materialtype"
					set rs_mattype=conn.execute(sql)
					%>		  		
					<select name="materialtype" readonly>			
						<%
							do while rs_mattype.eof=false
						%>
							<option value=<%=rs_mattype("materialtype")%> <%if rs_mattype("materialtype")=rs("materialtype") then%>selected="selected"<%end if%>><%=rs_mattype("materialtype")%></option>
						<%
							rs_mattype.movenext
							loop
							rs_mattype.close
						%>		 
					</select>							
		  		&nbsp;&nbsp;&nbsp;&nbsp;代理公司
					<%
					sql="select * from agent order by company"
					set rs_agent=conn.execute(sql)
					%>		  		
					<select name="agent" readonly>			
						<%
							do while rs_agent.eof=false
						%>
							<option value="<%=rs_agent("company")%>" <%if rs_agent("company")=rs("agent") then%>selected="selected"<%end if%>><%=rs_agent("company")%></option>
						<%
							rs_agent.movenext
							loop
							rs_agent.close
						%>								
					</select>		  	
		  		&nbsp;&nbsp;&nbsp;&nbsp;进口批文
					<%
					sql="select * from piwen order by company"
					set rs_piwen=conn.execute(sql)
					%>		  		
					<select name="piwen" readonly>			
						<%
							do while rs_piwen.eof=false
						%>
							<option value="<%=rs_piwen("company")%>" <%if rs_piwen("company")=rs("piwen") then%>selected="selected"<%end if%>><%=rs_piwen("company")%></option>
						<%
							rs_piwen.movenext
							loop
							rs_piwen.close
						%>								
					</select>			
					&nbsp;&nbsp;&nbsp;两证
		  		<input name="twodocumentsready" id="twodocumentsready" value="<%=rs("twodocumentready")%>" class="datepicker"  style="width:80px;border:none" readonly>										 		
				</td>
      </tr>    
      <tr>	  
	    	<td align="right" height="30">港口：</td>
        <td class="category">
					<%
					sql="select * from port"
					set rs_port=conn.execute(sql)
					%>	        
					<select name="destination" readonly>			
						<%
							do while rs_port.eof=false
						%>
							<option value=<%=rs_port("port")%> <%if rs_port("port")=rs("destination") then%>selected="selected"<%end if%>><%=rs_port("port")%></option>
						<%
							rs_port.movenext
							loop
							rs_port.close
						%>											
					</select>		
					&nbsp;&nbsp;&nbsp;&nbsp;靠泊码头
					<%
					sql="select * from terminal"
					set rs_terminal=conn.execute(sql)
					%>						
					<select name="terminal" readonly>		
						<%
							do while rs_terminal.eof=false
						%>
							<option value=<%=rs_terminal("terminal")%> <%if rs_terminal("terminal")=rs("terminal") then%>selected="selected"<%end if%>><%=rs_terminal("terminal")%></option>
						<%
							rs_terminal.movenext
							loop
							rs_terminal.close
						%>						
					</select>			
					&nbsp;&nbsp;&nbsp;&nbsp;船公司
					<%
					sql="select * from carrier"
					set rs_carrier=conn.execute(sql)
					%>							
					<select name="carrier" readonly>			
						<%
							do while rs_carrier.eof=false
						%>
							<option value=<%=rs_carrier("carrier")%> <%if rs_carrier("carrier")=rs("carrier") then%>selected="selected"<%end if%>><%=rs_carrier("carrier")%></option>
						<%
							rs_carrier.movenext 
							loop
							rs_carrier.close
						%>											
					</select>				
					&nbsp;&nbsp;&nbsp;&nbsp;船名航次
					<input name="shipname" style="width:150px;border:none" readonly maxlength="20">							
				</td>
      </tr>               
      <tr>	  
	    	<td align="right" height="30">动检证：</td>
        <td class="category">
		  		<input name="dongjian" readonly style="width:120px;border:none" value=<%=rs("dongjian")%>>
		  		<input name="dongjiancom" readonly style="border:none" value=<%=rs("dongjiancom")%>>
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">自动证：</td>
        <td class="category">
		  		<input name="zidong" readonly style="width:120px;border:none" value=<%=rs("zidong")%>>
		  		<input name="zidongcom" readonly style="border:none" value=<%=rs("zidongcom")%>>
		  		&nbsp;&nbsp;&nbsp;&nbsp;自动证交批日期
		  		<input name="zidongapplydate" id="zidongapplydate" class="datepicker"  style="width:80px;border:none" readonly value=<%=rs("zidongapplydate")%>>	  				  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;自动证上报日期
		  		<input name="zidongreportdate" id="zidongreportdate" class="datepicker"  style="width:80px;border:none" readonly value=<%=rs("zidongreportdate")%>>	  				  		
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">预保日：</td>
        <td class="category">
		  		<input name="planinsurance" id="planinsurance" class="datepicker" style="width:80px;border:none" readonly value=<%=rs("planinsurance")%>>
		  		&nbsp;&nbsp;&nbsp;&nbsp;补保日
		  		<input name="supinsurance" id="supinsurance" class="datepicker"  style="width:80px;border:none" readonly value=<%=rs("supinsurance")%>>
		  		&nbsp;&nbsp;&nbsp;&nbsp;保费支付日
		  		<input name="insurancepayment" id="insurancepayment" class="datepicker" style="width:80px;border:none" readonly value=<%=rs("insurancepayment")%>>	  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;保单
		  		<input name="insurancenumber" style="width:120px;border:none" readonly value=<%=rs("insurancenumber")%>>
				</td>
      </tr>      
      <tr>	  
	    	<td align="right" height="30">预计装船月份：</td>
        <td class="category">
		  		<input name="planship" style="width:100px;border:none" readonly>
		  		&nbsp;&nbsp;&nbsp;&nbsp;实际装船期
		  		<input name="shipdate" id="shipdate" class="datepicker" style="width:80px;border:none" readonly value=<%=rs("shipdate")%>>
		  		&nbsp;&nbsp;&nbsp;&nbsp;预计到港期
		  		<input name="boarddate" id="boarddate" class="datepicker" style="width:80px;border:none" readonly value=<%=rs("boarddate")%>>  		
		  		&nbsp;&nbsp;&nbsp;&nbsp;客户交货期
		  		<input name="plandeliverydate" style="width:100px;border:none" readonly value=<%=rs("plandeliverydate")%>>
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">箱号：</td>
        <td class="category">
		  		<input name="case" style="width:120px;border:none" readonly maxlength="15" value=<%=rs("case")%>>
		  		&nbsp;&nbsp;&nbsp;提单号
		  		<input name="ladnumber" style="width:120px;border:none" readonly maxlength="20" value=<%=rs("ladnumber")%>>
		  		&nbsp;&nbsp;&nbsp;卫生证号
		  		<input name="weishengzheng" style="width:120px;border:none" readonly maxlength="20" value=<%=rs("weishengzheng")%>>	
					&nbsp;&nbsp;&nbsp;铅封号
					<input name="locknumber" style="width:120px;border:none" readonly maxlength="20" value=<%=rs("locknumber")%>>		
		  		&nbsp;&nbsp;&nbsp;电子信息
					<select name="einformation" readonly>			
						  <option value="有" <%if rs("einformation")="True" then%>selected="selected"<%end if%>>有</option>		
						  <option value="无" <%if rs("einformation")="False" then%>selected="selected"<%end if%>>无</option>										
					</select>	
		  		&nbsp;&nbsp;&nbsp;清真证明
					<select name="MuslimCertification" readonly>			
						  <option value="有" <%if rs("MuslimCertification")="True" then%>selected="selected"<%end if%>>有</option>		
						  <option value="无" <%if rs("MuslimCertification")="False" then%>selected="selected"<%end if%>>无</option>										
					</select>									 		  		
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">预付款日期：</td>
        <td class="category">
		  		<input name="paydate" id="paydate" class="datepicker" style="width:80px;border:none" readonly value=<%=rs("prepaydate")%>>
		  		&nbsp;&nbsp;&nbsp;预付款金额
		  		<input name="prepayment" style="width:100px;border:none" readonly maxlength="15" value=<%=rs("prepayment")%> >
		  		&nbsp;&nbsp;&nbsp;币种
					<%
					sql="select * from trancurrency"
					set rs_currency=conn.execute(sql)
					%>
					<select name="currency" readonly>
						<%
							do while rs_currency.eof=false
						%>
							<option value="<%=rs_currency("currency")%>" <%if rs("trancurrency")=rs_currency("currency") then%>selected="selected"<%end if%>><%=rs_currency("currency")%></option>
						<%
							rs_currency.movenext
							loop
						%>
					</select>
				&nbsp;&nbsp;&nbsp;订金收到日期
				<input name="downpaymentdate" id="downpaymentdate" class="datepicker" style="width:80px;border:none" readonly value=<%=rs("downpaymentreceiptdate")%>>	
				&nbsp;&nbsp;&nbsp;成交条款
					<%
					sql="select * from tradingterm"
					set rs_term=conn.execute(sql)
					%>
					<select name="tradingterm" readonly>
						<%
							do while rs_term.eof=false
						%>
							<option value="<%=rs_term("tradingterm")%>" <%if rs("tradingterm")=rs_term("tradingterm") then%>selected="selected"<%end if%>><%=rs_term("tradingterm")%></option>
						<%
							rs_term.movenext
							loop
						%>
					</select>		
				</td>
      </tr> 
      <tr>	  
	    	<td align="right" height="30">尾款支付日期：</td>
        <td class="category">
		  		<input name="finalpaymentdate" id="finalpaymentdate" class="datepicker" style="width:80px;border:none" readonly value=<%=rs("finalpaymentdate")%>>
		  		&nbsp;&nbsp;&nbsp;免箱期
		  		<input name="freestayperiod" style="width:100px;border:none" readonly value=<%=rs("freestayperiod")%>>
		  		&nbsp;&nbsp;&nbsp;卫生证版本 
				<input name="weishengzhengversion" style="width:100px;border:none" readonly value=<%=rs("weishengzhengversion")%>>
				&nbsp;&nbsp;&nbsp;放行日期
		  		<input name="passdate" id="passdate" class="datepicker" style="width:80px;border:none" readonly value=<%=rs("passdate")%>>
				</td>
      </tr>       
      <tr>	  
	    	<td align="right" height="30">到单日期：</td>
        <td class="category">
		  		<input name="documentarrival" id="documentarrival" class="datepicker" style="width:80px;border:none" readonly value=<%=rs("documentarrival")%>>
		  		&nbsp;&nbsp;&nbsp;预计赎单日期
		  		<input name="planretirebill" id="planretirebill" class="datepicker" style="width:80px;border:none" readonly value=<%=rs("planretirebill")%>>	  		
		  		&nbsp;&nbsp;&nbsp;我司赎单日期
		  		<input name="internalretirebill" id="internalretirebill" class="datepicker" style="width:80px;border:none" readonly value=<%=rs("internalretirebill")%>>		  		
		  		&nbsp;&nbsp;&nbsp;代理赎单日期
		  		<input name="externalretirebill" id="externalretirebill" class="datepicker" style="width:80px;border:none" readonly value=<%=rs("externalretirebill")%>>		  		
				</td>
      </tr>       
      <tr>	  
	    	<td align="right" height="30">送货日期：</td>
        <td class="category">
		  		<input name="cargodate" id="cargodate" class="datepicker" style="width:80px;border:none" readonly value=<%=rs("cargodate")%>>
		  		&nbsp;&nbsp;&nbsp;送货方向
		  		<input name="cargodir" style="width:150px;border:none" readonly value=<%=rs("cargodir")%>>
		  		&nbsp;&nbsp;&nbsp;交单日期
		  		<input name="deliverydate" id="deliverydate" class="datepicker" style="width:80px;border:none" readonly value=<%=rs("deliverydate")%>>	  		
				</td>
      </tr> 
      <tr>
        <td align="right" height="30" valign="top">单证问题：</td>
        <td class="category" valign="top">
        	<textarea name="documentissue" cols="30" rows="3" value=<%=rs("documentissue")%> readonly></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="documentcheck" style="vertical-align:top"> 审单情况 </label>
        	<textarea name="documentcheck" id="documentcheck" cols="30" rows="3" value=<%=rs("documentcheck")%> readonly></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="documentdelivery" style="vertical-align:top"> 寄单情况 </label>
        	<textarea name="documentdelivery" id="documentdelivery" cols="30" rows="3" value=<%=rs("documentdelivery")%> readonly></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="clearcustom" style="vertical-align:top">清关进程</label>
        	<textarea name="clearcustom" id="clearcustom" cols="30" rows="3" value=<%=rs("clearcustom")%> readonly></textarea>            	   	
        </td>
      </tr>	 
      <tr>
        <td align="right" height="30" valign="top">保证金情况：</td>
        <td class="category" valign="top">
        	<textarea name="warrantyinformation" cols="30" rows="3" value=<%=rs("warrantyinformation")%> readonly></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="warrantystatus" style="vertical-align:top">保证金进度</label>
					<%
					sql="select * from warrantystatus"
					set rs_warrstatus=conn.execute(sql)
					%>
					<select name="warrantystatus"  style="vertical-align:top" readonly>
						<%
							do while rs_warrstatus.eof=false
						%>
							<option value="<%=rs_warrstatus("warrantystatus")%>" <%if rs("warrantystatus")=rs_warrstatus("warrantystatus") then%>selected="selected"<%end if%>><%=rs_warrstatus("warrantystatus")%></option>
						<%
							rs_warrstatus.movenext
							loop
						%>
					</select>	 
					<label for="warranty" style="vertical-align:top">&nbsp;&nbsp;&nbsp;保证金金额</label>	
					<input name="warranty" style="vertical-align:top;width:100px;border:none" readonly maxlength="15" value=<%=rs("warranty")%> >
					<label for="returnwarranty" style="vertical-align:top">&nbsp;&nbsp;&nbsp;退保证金金额</label>	
					<input name="returnwarranty" style="vertical-align:top;width:100px;border:none" readonly maxlength="15" value=<%=rs("returnwarranty")%> >
        </td>
      </tr>	        
      <tr>
        <td align="right" height="30" valign="top">少货情况：</td>
        <td class="category" valign="top">
        	<textarea name="stockmiss" cols="30" rows="3" value=<%=rs("stockmiss")%> readonly></textarea>&nbsp;&nbsp;&nbsp;
        	<label for="claiminformation" style="vertical-align:top">索赔情况</label>
        	<textarea name="claiminformation" id="claiminformation" cols="30" rows="3" value=<%=rs("claiminformation")%> readonly></textarea>&nbsp;&nbsp;&nbsp;
					<label for="claimstatus" style="vertical-align:top">索赔状态</label>
					<%
					sql="select * from claimstatus"
					set rs_claimstatus=conn.execute(sql)
					%>
					<select name="claimstatus"  style="vertical-align:top" readonly>
						<%
							do while rs_claimstatus.eof=false
						%>
							<option value="<%=rs_claimstatus("claimstatus")%>" <%if rs("claimstatus")=rs_claimstatus("claimstatus") then%>selected="selected"<%end if%>><%=rs_claimstatus("claimstatus")%></option>
						<%
							rs_claimstatus.movenext
							loop
						%>
					</select>					
					<label for="claimprocessor" style="vertical-align:top">索赔负责人</label>
					<select name="claimprocessor" style="vertical-align:top" readonly>
	    			<% for i = 0 to customCount-1 %>
	 						<option value=<%=custom(i)%> <%if rs("claimprocessor")=custom(i) then%>selected="selected"<%end if%>><%=custom(i)%></option>
	 					<% next %>  			
					</select>&nbsp;&nbsp;&nbsp;								
        </td>
      </tr>	       
      <tr>
            <td height="30">&nbsp;</td>
              <td class="category">
            <input type="button" value=" 返回 " onClick="window.history.go(-1);" class="button">&nbsp;&nbsp;&nbsp;&nbsp;        
            </td>
      </tr>  
</table>  
</form>
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
	  <form name="formitem">
      <tr>
		<input type="button" value="增加一行" class="button" disabled />&nbsp;&nbsp;&nbsp;
		<input type="button" value="删除一行" class="button" disabled />
      </tr>	      
      <tr>
		<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1" id="item">
		  <tr>
		  	<td width="20"></td>			
	        <td width="30" align="center">项目号</td>
	        <td width="100" align="center">品名</td>
	        <td width="100" align="center" >客户</td> 	
	        <td width=100" align="center" >规格</td> 		
	        <td width=80" align="center" >合同重量</td> 			
	        <td width="40" align="center" >重量单位</td> 		
	        <td width="80" align="center" >实际净重</td> 				    
	        <td width="80" align="center" >采购价格</td> 			
	        <td width="40" align="center" >价格单位</td> 				        	            		        	        		        
	        <td width="50" align="center" >生产日期</td> 		
	        <td width="50" align="center" >箱数</td> 		
	        <td width="100" align="center" >发票总金额</td> 		
	        <td width="50" align="center" >币种</td> 		
	        <td width="100" align="center" >尾款金额</td> 					        			        			        			        			        
     	  </tr>		
     	  <%
     	  	sql = "select * from shipmentitem where shipmentnum="&request("shipment")
     	  	set rs_items=conn.execute(sql)
		    if not rs_items.eof then
		    	do while rs_items.eof=false     	  	
     	  %>	
		  <tr>
		  	<td width="20"><input type='checkbox' name='chkArr'  id='chkArr' readonly /></td>			
	        <td width="30" align="center" ><input type='text' name='itemno' id='itemno' value='<%=rs_items("itemnum")%>' style='width:30px' readonly/></td>
	        <td width="100" align="center" >
	        	<select name="material<%=rs_items("itemnum")%>" id="material<%=rs_items("itemnum")%>" style='width:100px' readonly>
	    			<% for i = 0 to materialCount-1 %>
	 					<option value='<%=material(i)%>' <%if rs_items("material")=material(i) then%>selected="selected"<%end if%>><%=material(i)%></option>
	 				<% next %>
	 			</select>
	        </td>
	        <td width="100" align="center" >
	        	<select name="customer<%=rs_items("itemnum")%>" id="customer<%=rs_items("itemnum")%>" style='width:100px' readonly>
	    			<% for i = 0 to customerCount-1 %>
	 					<option value='<%=customer(i)%>' <%if rs_items("customer")=customer(i) then%>selected="selected"<%end if%>><%=customer(i)%></option>
	 				<% next %>
	 			</select>	        	
	        </td> 	
	        <td width=100" align="center" ><input type='text' name="spec<%=rs_items("itemnum")%>" id="spec<%=rs_items("itemnum")%>" value='<%=rs_items("spec")%>' style='width:100px' readonly/></td> 		
	        <td width=80" align="center" ><input type='text' name="contractweight<%=rs_items("itemnum")%>" id="contractweight<%=rs_items("itemnum")%>" value='<%=rs_items("contractweight")%>' style='width:80px' readonly/></td> 
	        <td width="40" align="center" >公斤</td> 		
	        <td width="80" align="center" ><input type='text' name="actualweight<%=rs_items("itemnum")%>" id="actualweight<%=rs_items("itemnum")%>" value='<%=rs_items("actualnetweight")%>' style='width:80px' readonly/></td>
	        <td width="80" align="center" ><input type='text' name="purchaseprice<%=rs_items("itemnum")%>" id="purchaseprice<%=rs_items("itemnum")%>" value='<%=rs_items("purchaseprice")%>' style='width:80px' readonly/></td> 	
	        <td width="40" align="center" >公斤</td> 				        	            		        	        		        
	        <td width="50" align="center" ><input name="producedate<%=rs_items("itemnum")%>" id="producedate<%=rs_items("itemnum")%>" value="<%=rs_items("productiondate")%>" style='width:80px' class="datepicker" readonly/></td> 		
	        <td width="50" align="center" ><input type='text' name="casenum<%=rs_items("itemnum")%>" id="casenum<%=rs_items("itemnum")%>" value='<%=rs_items("casenumber")%>' style='width:80px' readonly/></td> 		
	        <td width="100" align="center" ><input type='text' name="invamount<%=rs_items("itemnum")%>" id="invamount<%=rs_items("itemnum")%>" value='<%=rs_items("invoiceamount")%>' style='width:80px' readonly/></td> 	
	        <td width="50" align="center" >
					<%
					sql="select * from trancurrency"
					set rs_currency=conn.execute(sql)
					%>
					<select name="invcurr<%=rs_items("itemnum")%>" style='width:50px' readonly>
						<%
							do while rs_currency.eof=false
						%>
							<option value="<%=rs_currency("currency")%>" <%if rs_items("invoicecurrency")=rs_currency("currency") then%>selected="selected"<%end if%>><%=rs_currency("currency")%></option>
						<%
							rs_currency.movenext
							loop
						%>
					</select>	        	
	        </td> 		
	        <td width="100" align="center" ><input type='text' name="finalamount<%=rs_items("itemnum")%>" id="finalamount<%=rs_items("itemnum")%>" value='<%=rs_items("finalpayment")%>' style='width:80px' readonly/></td>
     	  </tr>	
		  <%
		    	rs_items.movenext
		  		loop
		  	end if
		  %>     	  
        </table>			  	
      </tr>	          

  </form>
</table>
</form>
</td>
<td></td>
</tr> 
</table>

</body>
</html>