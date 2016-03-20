<!--#include file="include/SQLConn.inc" -->
<%

  ID = Request("ID")

  If ID <> "" Then


     sql = "Select * from CouponType where ID = '"& ID &"'"

     Set rs = Conn.Execute(sql)


     CouponType = rs("CouponType")

     ProductType = rs("ProductType")

   End If

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
 
 if (document.fm1.CouponType.value == "")
  {
   alert("Please input coupon type.");
   document.fm1.CouponType.focus();
   return false;
  }

document.fm1.submit();
}



//-->
</SCRIPT>
</head>

<body leftmargin="0" topmargin="0" >
<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
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


<form name="fm1" method="post" action="">

  <table width="99%" border="0" class="normal">
    <tr> 
      <td colspan="2" class="BlueClr"><font size="5" face="Times New Roman, Times, serif">Update price </font></td>
    </tr>

    <tr> 
      <td colspan="2" bordercolor="#999966" class="ValidHead"></td>
    </tr>

 <tr><td>Coupon Type: </td>
      <td>
<input name="CouponType" type=text value="<% = Coupontype %>" size="5" maxlength="3">	    
</td>
   </tr>
 <tr>
<td>Product Type:</td>
<td>
<select size="1" name="ProductType" class="common">
			<option value="54" <%if ProductType = "54" then%>Selected<%End If%>>V 能量</option>
			<option value="53" <%if ProductType = "53" then%>Selected<%End If%>>無鉛</option>
			<option value="52" <%if ProductType = "52" then%>Selected<%End If%>>柴油</option>
			<option value="55" <%if ProductType = "55" then%>Selected<%End If%>>石油氣</option>
            <option value="CS" <%if ProductType = "CS" then%>Selected<%End If%>>便利店</option>
            <option value="LB" <%if ProductType = "LB" then%>Selected<%End If%>>機油</option>
            <option value="CW" <%if ProductType = "CW" then%>Selected<%End If%>>洗車</option>
	</select> 	
  
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

<%  If ProductType <> "" Then %>
<input type=hidden name=whatdo value='edit_cp'>
<% Else %>
<input type=hidden name=whatdo value='add_cp'>
<%  end if %>

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