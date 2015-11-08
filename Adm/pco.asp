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
 
 if (document.fm1.Face_Value.value == "")
  {
   alert("Please input face value.");
   document.fm1.Face_Value.focus();
   return false;
  }
if (document.fm1.Product_Type.value == "")
  {
   alert("Please input coupon type.");
   document.fm1.Product_Type.focus();
   return false;
  }

document.fm1.submit();
}




function dosave(){
document.fm1.action="hsemis.asp?page_id=execute";
  if (document.fm1.PCDate.value == "")
  {
   alert("Please enter the date.");
   document.fm1.PCDate.focus();
   return false;
  }

document.fm1.whatdo.value="modify_cm";
document.fm1.submit();
}

function selectdate(){

  if (document.fm1.priority.value == "1")
   {
    document.fm1.PCDate.value = document.fm1.select_urgent.value;
    return true;

   }

 if (document.fm1.priority.value == "2")
   {
    document.fm1.PCDate.value = document.fm1.select_normal.value;
    return true;

   }
if (document.fm1.priority.value == "3")
   {
    document.fm1.PCDate.value = document.fm1.select_nextmonth.value;
    return true;

   }
}

function newWindow(file,window) {
   var URL = document.fm1.treeid.options[document.fm1.treeid.selectedIndex].value;
    URL = 'lstlocation.asp?treeid='  + URL
  msgWindow=open(URL,window,'scrollbars=yes, resizable=yes,width=800,height=600');
  if(msgWindow.opener == null) msgWindow.opener = self;
}

function dochange(){
 document.fm1.action="cmo.asp";
 document.fm1.submit();
}

//-->
</SCRIPT>
</head>

<body leftmargin="0" topmargin="0" >
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
      <td colspan="2" class="BlueClr"><font size="5" face="Times New Roman, Times, serif">Create Pro Coupon </font></td>
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

 <tr><td>Face Value: </td>
      <td>
<input name="Face_Value" type=text value="">	    
</td>
   </tr>
 <tr>
<td>Type:</td>
<td>
<input name="Product_Type" type=text value="">	    
</td>
</tr>
     
 <tr><td>Batch: </td>
      <td>
<input name="Batch" type=text value="">	    
</td>
   </tr>

<tr><td>Start Range:</td>
      <td>
<input name="Start_Range" type=text value="">	    
</td>
   </tr>
<tr><td>End Range:</td>
      <td>
<input name="End_Range" type=text value="">	    
</td>
   </tr>

     <tr>
	<td><b>Coupon Value</b></td>
      <td>

</td>
   	</tr>

<tr><td>Category:</td>
      <td>
<input name="Category" type=text value="">	    
</td>
   </tr>
<tr><td>Dealer A/C:</td>
      <td>
<input name="Dealer_AC" type=text value="">	    
</td>
   </tr>
<tr><td>Canopy Disc:</td>
      <td>
<input name="Canopy_Disc" type=text value="">	    
</td>
   </tr>
<tr><td>Expiry Date: (dd-mm-yyyy)</td>
      <td>
<input name="Expiry_Date" type=text value="">	    
</td>
   </tr>
<tr><td>Issue Date: (dd-mm-yyyy)</td>
      <td>
<input name="Issue_date" type=text value="">	    
</td>
   </tr>
<tr><td>Completed:</td>
      <td>
<input name="Completed" type=text value="">	    
</td>
   </tr>
<tr><td>Excel Type:</td>
      <td>
<input name="Excel_Type" type=text value="">	    
</td>
   </tr>
<tr><td>Effective Date:</td>
      <td>
<input name="Effective_Date" type=text value="">	    
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
<input type=hidden name=whatdo value='add_pco'>

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