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


%>
<HTML>
<HEAD>
	<link rel="stylesheet" type="text/css" href="include/publish.css" />

<TITLE>Check Digit</TITLE>

<!-- Load the javascript code -->

<SCRIPT language=JavaScript>
<!--

function CheckDigit() {

key = document.barCodeForm.keynumber.value
bc  = document.barCodeForm.Barcode.value
checkdigit = bc.substr(3,1);
bc = bc.replace(/[^0-9]+/g,'');

//document.barCodeForm.Barcode.value = bc;
total = 0;

if (bc.length=16){
bc = bc.substr(0,15)
}

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



//total = total + key;
//Determine the checksum. if modDigit = 10, modDigit = 0
modDigit  = (10 - ((total + 17) % 10)) % 10;
modDigit2 = (10 - ((total + 23) % 10)) % 10;


document.barCodeForm.checkdigit.value = modDigit;
document.barCodeForm.checkdigit2.value = modDigit2;
document.barCodeForm.Barcode2.value = bc.concat(modDigit2);

}

//-->
</script>


</HEAD>

<BODY onload="document.barCodeForm.Barcode.focus();">

<h1>Check Digit Generator</h1>

<div align="center">

<FORM NAME="barCodeForm"  method="post"  action="Checkdigit.asp">

<table width="80%" border="1"  cellpadding="8" cellspacing="0" class="Report">
										
<tr>

<td width="30%">
First Key Number
</td>

<td class="common"> 

<INPUT name="keynumber" size="2" value="17">

		
</td></tr>

<tr>

<td width="30%">
Second Key Number
</td>

<td class="common"> 

<INPUT name="keynumber2" size="2" value="23">

		
</td></tr>
			
<tr>

<td>
Coupon Number
</td>

<td class="common"> 

<INPUT name="Barcode" size="20" value="">

		
</td></tr>
								
	 
							
								
								<tr><td class="common">
								
			
								
								
		
							</td><td class="common">
							<INPUT TYPE=button onClick="CheckDigit();" value="Generate">
							</td></tr>
		
									<tr><td class="common">
				First Check Digit

								</td><td class="common">
				<INPUT name="checkdigit" size="2" value="">

								</td></tr>

<tr><td class="common">
				Second Check Digit

								</td><td class="common">
				<INPUT name="checkdigit2" size="2" value="">

								</td></tr>
							
								
<tr>

<td>
New 16 digits Coupon Number
</td>

<td class="common"> 

<INPUT name="Barcode2" size="20" value="">

		
</td></tr>

								</table>
								</FORM>
							

					
</div>
	

</BODY>
</HTML>