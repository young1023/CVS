<!--#include file="include/SQLConn.inc" -->
<!--#include file ="js/OVERLIB.JS" -->
<!--#include file ="js/OVERLIB_MINI.JS" -->
<!--#include file ="js/select_date.JS" -->

<%

flag = trim(request.form("whatdo"))
RangeID        = Request("ID")

If RangeID <> "" Then

   sql = "Select * from Coupon_Redempted where RangeID = "&RangeID

   Set rs = conn.execute(sql)

  Face_Value = trim(rs("FaceValue"))
 
  Product_Type = replace(trim(rs("Product_Type")),"'","''")
 
  Batch = replace(trim(rs("Batch")),"'","''")
 
  Start_Range = replace(trim(rs("Start_Range")),"'","''")
 
  End_Range = trim(rs("End_Range"))

 


End if

If flag = "add_range" then

   
  If RangeID <> "" Then

    ' Delete record

    sql1 = "Delete From Coupon_Redempted where RangeID = "&RangeID

    Conn.execute(sql1)


  End If
 
  Face_Value = replace(trim(request.form("Face_Value")),"'","''")
  
  Product_Type = replace(trim(request.form("Product_Type")),"'","''")
 
  Batch = replace(trim(request.form("Batch")),"'","''")
 
  Start_Range = replace(trim(request.form("Start_Range")),"'","''")
 
  End_Range = request.form("End_Range")

 
      sql = "Insert into Coupon_Redempted (FaceValue, Product_Type, Batch, Start_Range, End_Range)"
  
      sql = sql & " values('"&Face_Value&"', '"&Product_Type&"', '"&Batch&"', '"&Start_Range&"', '"&End_Range&"')"
  
   response.write sql
  'response.end
     conn.execute(sql)
  
     response.redirect "coupon_redempted1.asp"


  end if



%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>Shell CVS Administration</title>
<link rel="stylesheet" type="text/css" href="include/hse.css" />
<SCRIPT language=JavaScript>
<!--
function dosubmit(){
 document.fm1.action="Analy_NewRange1.asp";
 
 if (document.fm1.Face_Value.value == "")
  {
   alert("Please input face value.");
   document.fm1.Face_Value.focus();
   return false;
  }
var x = document.fm1.Face_Value.value
 if (isNaN(x)) 
  {
    alert("Please input numbers.");
    return false;
  }
if (x.length!=3) 
  {
    alert("Face Value should be 3 digits.");
    document.fm1.Face_Value.focus();
    return false;
  }
if (document.fm1.Product_Type.value == "")
  {
   alert("Please input coupon type.");
   document.fm1.Product_Type.focus();
   return false;
  }
var y = document.fm1.Product_Type.value
 if (isNaN(y)) 
  {
    alert("Please input numbers.");
    return false;
  }
if (document.fm1.Batch.value == "")
  {
   alert("Please input Batch.");
   document.fm1.Batch.focus();
   return false;
  }
var z = document.fm1.Batch.value
 if (isNaN(z)) 
  {
    alert("Please input numbers.");
    return false;
  }
if (document.fm1.Start_Range.value == "")
  {
   alert("Please input Start Range.");
   document.fm1.Start_Range.focus();
   return false;
  }
var s = document.fm1.Start_Range.value
 if (isNaN(s)) 
  {
    alert("Please input numbers.");
    return false;
  }
if (document.fm1.End_Range.value == "")
  {
   alert("Please input End Range.");
   document.fm1.End_Range.focus();
   return false;
  }
var e = document.fm1.End_Range.value
 if (isNaN(e)) 
  {
    alert("Please input numbers.");
    return false;
  }

document.fm1.whatdo.value="add_range";

document.fm1.submit();
}


//-->
</SCRIPT>

<script language="JavaScript" src="include\ts_picker.js">
</script>
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
      <td colspan="2" class="BlueClr">Pro Coupon Range</td>
    </tr>
    <tr> 
      <td colspan="2"></td>
    </tr>

    <tr> 
      <td colspan="2" bordercolor="#999966" class="ValidHead"></td>
    </tr>

 <tr><td width="30%">Face Value: </td>
      <td>
<input name="Face_Value" type=text value="<% = Face_Value %>" maxlength="3">	    
</td>
   </tr>
 <tr>
<td>Type:</td>
<td>
<input name="Product_Type" type=text value="<% = Product_Type %>" maxlength="2">	    
</td>
</tr>
     
 <tr><td>Batch: </td>
      <td>
<input name="Batch" type=text value="<% = Batch %>" maxlength="3">	    
</td>
   </tr>

<tr><td>Start Range:</td>
      <td>
<input name="Start_Range" type=text value="<% = Start_Range %>" maxlength="6">	    
</td>
   </tr>
<tr><td>End Range:</td>
      <td>
<input name="End_Range" type=text value="<% = End_Range %>" maxlength="6">	    
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
<input type=hidden name=whatdo value="">
<input type=hidden name="ID" value="<% = RangeID %>">
<input type="Reset" value="  Reset  " class="common">


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