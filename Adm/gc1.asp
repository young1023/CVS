<!--#include file="include/SQLConn.inc" -->
<% 


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>Shell CVS Administration</title>
<link rel="stylesheet" type="text/css" href="include/hse.css" />
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

range = parseInt(er) - parseInt(sr) + 1

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
</SCRIPT>
</head>

<body leftmargin="0" topmargin="0" onload="document.barCodeForm.facevalue.focus();">
<!--#include file="include/header.inc" -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="3" class="NavyBlue" height="3"></td>
  </tr>
  <tr>
    <td colspan="3" class="NavyBlue">
    <table border="0" width="100%" cellspacing="0" cellpadding="0" class="TitleBar">
      <tr>
        <td width="100%">
        <table border="0" width="100%" cellspacing="0" cellpadding="0" class="TitleBar">
            <tr>
              <td><div align="center"><font class="Head" style="font-size: 13px">Shell CVS Administration</font></div></td>
            </tr>
            </table>
          </td>
          <td nowrap="true">
          </td>
        </tr>
      </table>
      </td>
    </tr>
    <tr>
      <td colspan="3" class="NavyBlue" height="1"></td>
    </tr>
    <tr>
      <td colspan="3" height="1"></td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="60%">
    <tr>
      <td width="1" class="HSEBlue"></td>
      <td valign="top">
      <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="4" class="HSEBlue"></td>
        </tr>
        <tr valign="top">
          <td height="25"><img src="images/Curve.gif" width="22" height="16" /></td>
        </tr>
        <tr valign="top">
          <td height="100%" align="middle">


 <table width="100%" border="0" cellpadding=0 cellspacing="0" bgcolor="#FFFFFF" height="100%">
                <tr>
                  <td valign="top" align="center" bgcolor="#E6EBEF">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E6EBEF" class="normal">
                     <tr> 
                          
                        <td height="28" align="center"><font color="#FF6600"><b>
	Coupon Generator</b></font></td>
                        </tr>
                        <tr> 
                          <td valign="top" align="center">
                          


<div align="center">

<FORM NAME="barCodeForm"  method="post"  action="GenCoupon.asp">

<table width="100%" border="0"  cellpadding="4" cellspacing="2" class="Report">
										
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

<INPUT name="facevalue" type="text" size="3" value="" MaxLength="3">
	
</td>

<td>
Coupon type 
</td>

<td class="common"> 

<INPUT name="coupontype" type="text" size="2" value="" MaxLength="2">
	
</td>

</tr>
<tr>

<td>
Batch
</td>

<td class="common" colspan="3"> 

<INPUT name="batch" size="3" type="text" value="" MaxLength="3">
	
</td>

</tr>

<tr>

<td>
Start Range
</td>

<td class="common"> 

<INPUT name="startrange" size="6" type="text" value="" MaxLength="6">
	
</td>

<td>
End Range
</td>

<td class="common"> 

<INPUT name="endrange" size="6" type="text" value="" MaxLength="6">
	
</td>

</tr>
								
	 
							
								
								<tr><td class="common">
								
			
								
								
		
							</td><td class="common" >
							<INPUT TYPE=button onClick="GenCoupon();" value="Generate">
	
							</td>


<td colspan="2">
	
</td>


</tr>
		
	
							
								
<tr>

<td>
16 digits Coupon Number
</td>

<td class="common" colspan="3"> 

<textarea rows="40" cols="80" id="Barcode3" onfocus="this.select();" onmouseup="return false;">
</textarea>
		
</td></tr>

								</table>
								</FORM>
							

					
</div>



<%
'-----------------------------------------------------------------------------
'
'      End of the main Content 
'
'-----------------------------------------------------------------------------
%>
</td>
              </tr>
                </table>
                </td>
                </tr>
              </table>
           
              </body>

              </html>