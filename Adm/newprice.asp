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
function dosubmit(){
 document.fm1.action="execute1.asp";
 
 if (document.fm1.product.value == "")
  {
   alert("Please input product type.");
   document.fm1.product.focus();
   return false;
  }
if (document.fm1.price.value == "")
  {
   alert("Please input unit price.");
   document.fm1.price.focus();
   return false;
  }

document.fm1.submit();
}



//-->
</SCRIPT>
</head>

<body leftmargin="0" topmargin="0" >
<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="60"><IMG style="LEFT: 0px; POSITION: absolute; TOP: 0px"  alt ="Shell logo" src="IMAGES/Shell.gif" border=0 ></IMG></td>
    <td align="right"></td>
  </tr>
</table>
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
      <td width="120" valign="top">
      <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" class="Common">
        <tr valign="top" align="center">
          <td class="HSEBlue" height="21"></td>
        </tr>
        <tr valign="top" align="center">
          <td>
              <br/>
              <p>
              <a href="pco.asp" title="Create new pro coupon" style="TEXT-DECORATION: none">
              New Pro Coupon<br>
              </a>
              </p>
            
              <p>
              <a href="pco_d.asp" title="pro coupon" style="TEXT-DECORATION: none">
              Pro Coupon<br>
              </a>
              </p>

 <p>
              <a href="master.asp" title="pro coupon" style="TEXT-DECORATION: none">
              Master List<br>
              </a>
              </p>

 <p>
              <a href="station.asp" title="pro coupon" style="TEXT-DECORATION: none">
              Station<br>
              </a>
              </p>

 <p>
              <a href="price.asp" title="pro coupon" style="TEXT-DECORATION: none">
              Price<br>
              </a>
              </p>
          </td>
        </tr>
      </table>
      </td>
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


<form name="fm1" method="post" action="">

  <table width="99%" border="0" class="normal">
    <tr> 
      <td colspan="2" class="BlueClr"><font size="5" face="Times New Roman, Times, serif">Update price </font></td>
    </tr>
    <tr> 
      <td width="35%"> </td>
      <td width="65%" align="right">Today is <% = submit_date %></td>
    </tr>
    <tr> 
      <td colspan="2"></td>
    </tr>
  <tr> 
      <td colspan="2"  align="right">
<font color="red">*</font>Denotes a mandatory field</td>
    </tr>
    <tr> 
      <td colspan="2" bordercolor="#999966" class="ValidHead"></td>
    </tr>

 <tr><td>Product: </td>
      <td>
<input name="product" type=text value="">	    
</td>
   </tr>
 <tr>
<td>Unit Price:</td>
<td>
<input name="price" type=text value="">	    
</td>
</tr>
     
 <tr><td>Effective Date: </td>
      <td>
<input name="EffectiveDate" type=text value="">	    
</td>
   </tr>

<tr><td>Expiry Date:</td>
      <td>
<input name="ExpiryDate" type=text value="">	    
</td>
   </tr>

	<tr>
<td></td>
<td>
  
</td>
	</tr>
     <tr>
 
    
                                        <td colspan="2" align="center"> 

<input type="button" value="  Submit  " onClick="dosubmit();" class="common">
<input type="Reset" value="  Reset  " class="common">
<input type=hidden name=whatdo value='add_pr'>

</td>
</tr>
</table>
</form>


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