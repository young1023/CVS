<% 
'on error resume next
'*********************************************************************************
'NAME       : Report.asp          
'DESCRIPTION: Check digit
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : Gary Yeung
'MODIFIED   : 
'********************************************************************************



<HTML>
<HEAD>
	<link rel="stylesheet" type="text/css" href="include/publish.css" />

<TITLE>Check Digit</TITLE>

<!-- Load the javascript code -->

<SCRIPT language=JavaScript>
<!--


// Check digit
function CheckDigit() {
//Get barCode field value, Strip all characters
//except numbers Repopulate field w/ new value
bc = document.barCodeForm.Barcode.value
barcode1 = document.barCodeForm.Barcode.value
checkdigit = bc.substr(3,1);
bc = bc.replace(/[^0-9]+/g,'');
document.barCodeForm.Barcode.value = bc;
total = 0;
//Get Even Numbers
for (i=bc.length-2; i>=0; i=i-2) {
if (i != 3){
tempeven = parseInt(bc.substr(i,1)) ;
if (tempeven > 9) {
tempeven = tempeven - 9;
}
total = total + tempeven ;
}
}
//Get Odd Numbers
for (i=bc.length-1; i>=0; i=i-2) {
temp = parseInt(bc.substr(i,1)) * 2 ;
if (temp > 9) {
temp = temp - 9;
}
total = total + temp;
}

//Determine the checksum. if modDigit = 10, modDigit = 0
modDigit = (10 - ((total + 17) % 10)) % 10;


if (checkdigit != modDigit){
alert("驗證失敗 - 禮券無效!");
document.barCodeForm.Barcode.focus();
return false;

}


//Populate the modDigit field with the value
//document.barCodeForm.barCodeShow.value = barcode1;
document.barCodeForm.action="execute.asp";
document.barCodeForm.submit();
}


//-->
</script>


</HEAD>

<BODY >

FORM NAME="barCodeForm"  method="post"  action="Checkdigit.asp">

								  <table width="99%" border="0" class="normal">
			
							
										<tr><td class="common"> 
											
											Enter client No., English Name or Chinese Name
										</td></tr>
										<tr><td class="common"> 
										<INPUT name="keyword" value="<%= Search_keyword %>">
											<INPUT TYPE=submit value="Filter">
											<INPUT TYPE=Reset value="Clear">

										</td></tr>

			

								
	 
								<tr><td class="common">
				
								</td></tr>
							
								
								<tr><td class="common">
								
			
								
								
		
							</td></tr>
							
							<tr><td class="common">
							<INPUT TYPE=button onClick="AssignValue();" value="Select Client">
							</td></tr>
		
								
								
								</table>
								</FORM>
							

					

	

</BODY>
</HTML>