<% 
'on error resume next
'*********************************************************************************
'NAME       : GenCoupon.asp          
'DESCRIPTION: Generate Coupon
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : Gary Yeung
'MODIFIED   : 
'********************************************************************************


%>
<!DOCTYPE html>
<HTML>
<HEAD>
	<link rel="stylesheet" type="text/css" href="include/publish.css" />

<TITLE>Generate Coupon</TITLE>

<!-- Load the javascript code -->

<SCRIPT language=JavaScript>
<!--

function GenCoupon() {

var range = 0;
var bc2 = "";

key = document.barCodeForm.keynumber.value
fv  = document.barCodeForm.facevalue.value
ct  = document.barCodeForm.coupontype.value
bt  = document.barCodeForm.batch.value
sr  = document.barCodeForm.startrange.value
er  = document.barCodeForm.endrange.value

range = parseInt(er) - parseInt(sr) 

// from start range to end range
for (j = 0; j < range; j++){

// combine all values to form a coupon number
bc  = fv + 0 + ct + bt + sr ;
bc = bc.replace(/[^0-9]+/g,'');

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
modDigit  = (10 - ((total + 17) % 10)) % 10;
modDigit2 = (10 - ((total + 23) % 10)) % 10;


bc1 =  fv + modDigit + ct + bt + sr + modDigit2 + "\n" ;

bc2 = bc2 + bc1

sr = parseInt(sr) + 1

sr = ZeroPadNumber(sr)


}

document.getElementById("Barcode3").value = bc2;

}

function ZeroPadNumber(nValue)
{
    if ( nValue < 10 )
    {
        return ( '00000' + nValue.toString () );
    }
    else if ( nValue < 100 )
    {
        return ( '0000' + nValue.toString () );
    }
    else if ( nValue < 1000 )
    {
        return ( '000' + nValue.toString () );
    }
      else if ( nValue < 10000 )
    {
        return ( '00' + nValue.toString () );
    }
    else if ( nValue < 100000 )
    {
        return ( '0' + nValue.toString () );
    }
    else
    {
        return ( nValue );
    }
}

//-->
</script>


</HEAD>

<BODY onload="document.barCodeForm.Barcode.focus();">

<h1>Coupon Generator</h1>

<div align="center">

<FORM NAME="barCodeForm"  method="post"  action="GenCoupon.asp">

<table width="90%" border="1"  cellpadding="4" cellspacing="2" class="Report">
										
<tr>

<td >
First Key Number
</td>

<td class="common"> 

<INPUT name="keynumber" type="text" size="2" value="17" readonly>
		
</td>

<td >
Second Key Number
</td>

<td class="common"> 

<INPUT name="keynumber2" type="text" size="2" value="23" readonly>

		
</td>



</tr>
			
<tr>

<td>
Face Value 
</td>

<td class="common"> 

<INPUT name="facevalue" type="text" size="3" value="045" MaxLength="3">
	
</td>

<td>
Coupon type 
</td>

<td class="common"> 

<INPUT name="coupontype" type="text" size="2" value="00" MaxLength="2">
	
</td>

</tr>
<tr>

<td>
Batch
</td>

<td class="common" colspan="3"> 

<INPUT name="batch" size="3" type="text" value="372" MaxLength="3">
	
</td>

</tr>

<tr>

<td>
Start Range
</td>

<td class="common"> 

<INPUT name="startrange" size="6" type="text" value="000001" MaxLength="6">
	
</td>

<td>
End Range
</td>

<td class="common"> 

<INPUT name="endrange" size="6" type="text" value="000010" MaxLength="6">
	
</td>

</tr>
								
	 
							
								
								<tr><td class="common">
								
			
								
								
		
							</td><td class="common" colspan="4">
							<INPUT TYPE=button onClick="GenCoupon();" value="Generate">
	
							</td></tr>
		
	
							
								
<tr>

<td>
New 16 digits Coupon Number
</td>

<td class="common" colspan="3"> 

<textarea rows="40" cols="130" id="Barcode3">
</textarea>
		
</td></tr>

								</table>
								</FORM>
							

					
</div>
	

</BODY>
</HTML>