<!--#include file="include/SQLConn.inc" -->
<!--#include file ="js/OVERLIB.JS" -->
<!--#include file ="js/OVERLIB_MINI.JS" -->
<!--#include file ="js/select_date.JS" -->

<%

flag = trim(request.form("whatdo"))
ID        = Request("ID")

If ID <> "" Then

   sql = "Select * from LeafLetBatch where ID = "&ID

   Set rs = conn.execute(sql)

  Face_Value1 = trim(rs("FaceValue1"))
 
  Coupon_Type1 = replace(trim(rs("Coupon_Type1")),"'","''")
 
  Batch1 = replace(trim(rs("Batch1")),"'","''")
 
  Start_Range1 = replace(trim(rs("Start_Range1")),"'","''")
 
  NoOfCoupon1 = trim(rs("NoOfCoupon1"))

  Face_Value2 = trim(rs("FaceValue2"))
 
  Coupon_Type2 = replace(trim(rs("Coupon_Type2")),"'","''")
 
  Batch2 = replace(trim(rs("Batch2")),"'","''")
 
  Start_Range2 = replace(trim(rs("Start_Range2")),"'","''")
 
  NoOfCoupon2 = trim(rs("NoOfCoupon2"))

  NoOfLeafLet = trim(rs("NoOfLeafLet"))
 


End if

if flag = "add_range" then
 
  Face_Value1 = replace(trim(request.form("Face_Value1")),"'","''")
  
  Coupon_Type1 = replace(trim(request.form("Coupon_Type1")),"'","''")
 
  Batch1 = replace(trim(request.form("Batch1")),"'","''")
 
  Start_Range1 = replace(trim(request.form("Start_Range1")),"'","''")

  NoOfCoupon1 = request.form("NoOfCoupon1")

  Face_Value2 = replace(trim(request.form("Face_Value2")),"'","''")
  
  Coupon_Type2 = replace(trim(request.form("Coupon_Type2")),"'","''")
 
  Batch2 = replace(trim(request.form("Batch2")),"'","''")
 
  Start_Range2 = replace(trim(request.form("Start_Range2")),"'","''")
 
  NoOfCoupon2 = request.form("NoOfCoupon2")

  NoOfLeaflet = request.form("NoOfLeaflet")

      
 
      sql = "Insert into LeafletBatch (FaceValue1, Coupon_Type1, Batch1, Start_Range1, NoOfCoupon1, "

      sql = sql & " FaceValue2, Coupon_Type2, Batch2, Start_Range2, NoOfCoupon2, NoOfLeaflet) "
  
      sql = sql & " values ('"&Face_Value1&"', '"&Coupon_Type1&"', '"&Batch1&"', '"&Start_Range1&"', "

      sql = sql & " '"&NoOfCoupon1&"', '"&Face_Value2&"', '"&Coupon_Type2&"', '"&Batch2&"', "

      sql = sql & " '"&Start_Range2&"', '"&NoOfCoupon2&"', '"&NoOfLeaflet&"')"
  
     'response.write sql
  
     conn.execute(sql)

       sql1 = "Select top 1 ID From LeafletBatch order by id desc"

       Set Rs1 = Conn.Execute(sql1)

       LeafLetID = Rs1("ID")
  
      for i=1 to NoOfLeafLet

        End_Range1 = Clng(Start_Range1) + Clng(NoOfCoupon1) - 1 

        End_Range1 = string(6 - LEN(End_Range1), "0") & End_Range1

        'response.write End_Range1

        sql2 = "Insert into Coupon_Redemption (LeafLetID, LeafLetNo, FaceValue, Coupon_Type, Batch, Start_Range, End_Range, NoOfCoupon)"

        sql2 = sql2 & " Values ("&LeafLetID&", "& i &", '"&Face_Value1&"', '"&Coupon_Type1&"', '"&Batch1&"', '"&Start_Range1&"', "

        sql2 = sql2 & " '"&End_Range1&"', "&NoOfCoupon1&")"

        'response.write sql2

        conn.execute(sql2)

        ' Increment of start range for next leaflet
        Start_Range1 = Clng(Start_Range1) + Clng(NoOfCoupon1)


        ' Increment of end range of 2nd batch for next leaflet
        End_Range2 = Clng(Start_Range2) + Clng(NoOfCoupon2) - 1

        End_Range2 = string(6 - LEN(End_Range1), "0") & End_Range2

       

        sql3 = "Insert into Coupon_Redemption (LeafLetID, LeafLetNo, FaceValue, Coupon_Type, Batch, Start_Range, End_Range, NoOfCoupon)"

        sql3 = sql3 & " Values ("&LeafLetID&",  "& i &", '"&Face_Value2&"', '"&Coupon_Type2&"', '"&Batch2&"', '"&Start_Range2&"', "

        sql3 = sql3 & " '"&End_Range2&"', "&NoOfCoupon2&" )"

        response.write sql3

        conn.execute(sql3)

        ' Increment of start range for next leaflet
        Start_Range2 = Clng(Start_Range2) + Clng(NoOfCoupon2)
        
      next



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
 
 if (document.fm1.Face_Value1.value == "")
  {
   alert("Please input face value at first batch.");
   document.fm1.Face_Value1.focus();
   return false;
  }
var x = document.fm1.Face_Value1.value
 if (isNaN(x)) 
  {
    alert("Please input numbers at first batch.");
    return false;
  }
if (x.length!=3) 
  {
    alert("Face Value at first batch should be 3 digits.");
    document.fm1.Face_Value1.focus();
    return false;
  }

if (document.fm1.Coupon_Type1.value == "")
  {
   alert("Please input coupon type at first batch.");
   document.fm1.Coupon_Type1.focus();
   return false;
  }
var y = document.fm1.Coupon_Type1.value
 if (isNaN(y)) 
  {
    alert("Please input numbers.");
    return false;
  }

if (y.length!=2) 
  {
    alert("Coupon type at first batch should be 2 digits.");
    document.fm1.Coupon_Type1.focus();
    return false;
  }


if (document.fm1.Batch1.value == "")
  {
   alert("Please input Batch.");
   document.fm1.Batch1.focus();
   return false;
  }
var z = document.fm1.Batch1.value
 if (isNaN(z)) 
  {
    alert("Please input numbers.");
    return false;
  }

if (z.length!=3) 
  {
    alert("Batch at first batch should be 3 digits.");
    document.fm1.Batch1.focus();
    return false;
  }


if (document.fm1.Start_Range1.value == "")
  {
   alert("Please input Start Range.");
   document.fm1.Start_Range1.focus();
   return false;
  }
var s = document.fm1.Start_Range1.value
 if (isNaN(s)) 
  {
    alert("Please input numbers.");
    return false;
  }

if (s.length!=6) 
  {
    alert("Start range at first batch should be 6 digits.");
    document.fm1.Start_Range1.focus();
    return false;
  }

if (document.fm1.NoOfCoupon1.value == "")
  {
   alert("Please input number of coupon at first batch.");
   document.fm1.NoOfCoupon1.focus();
   return false;
  }
var s = document.fm1.NoOfCoupon1.value
 if (isNaN(s)) 
  {
    alert("Please input number.");
    return false;
  }

if (document.fm1.Face_Value2.value == "")
  {
   alert("Please input face value at first batch.");
   document.fm1.Face_Value2.focus();
   return false;
  }
var x = document.fm1.Face_Value2.value
 if (isNaN(x)) 
  {
    alert("Please input numbers.");
    return false;
  }
if (x.length!=3) 
  {
    alert("Face Value at second batch should be 3 digits.");
    document.fm1.Face_Value2.focus();
    return false;
  }

if (document.fm1.Coupon_Type2.value == "")
  {
   alert("Please input coupon type at second batch.");
   document.fm1.Coupon_Type2.focus();
   return false;
  }
var y = document.fm1.Coupon_Type2.value
 if (isNaN(y)) 
  {
    alert("Please input numbers.");
    return false;
  }

if (y.length!=2) 
  {
    alert("Coupon type at second batch should be 2 digits.");
    document.fm1.Coupon_Type2.focus();
    return false;
  }


if (document.fm1.Batch2.value == "")
  {
   alert("Please input Batch.");
   document.fm1.Batch2.focus();
   return false;
  }
var z = document.fm1.Batch2.value
 if (isNaN(z)) 
  {
    alert("Please input numbers.");
    return false;
  }

if (z.length!=3) 
  {
    alert("Batch at first batch should be 3 digits.");
    document.fm1.Batch2.focus();
    return false;
  }


if (document.fm1.Start_Range2.value == "")
  {
   alert("Please input Start Range.");
   document.fm1.Start_Range2.focus();
   return false;
  }
var s = document.fm1.Start_Range2.value
 if (isNaN(s)) 
  {
    alert("Please input numbers.");
    return false;
  }

if (s.length!=6) 
  {
    alert("Start range at second batch should be 6 digits.");
    document.fm1.Start_Range2.focus();
    return false;
  }

if (document.fm1.NoOfCoupon2.value == "")
  {
   alert("Please input number of coupon at second batch.");
   document.fm1.NoOfCoupon2.focus();
   return false;
  }
var s = document.fm1.NoOfCoupon2.value
 if (isNaN(s)) 
  {
    alert("Please input number.");
    return false;
  }

if (document.fm1.NoOfLeaflet.value == "")
  {
   alert("Please input number of leaflet.");
   document.fm1.NNoOfLeaflet.focus();
   return false;
  }
var s = document.fm1.NoOfLeaflet.value
 if (isNaN(s)) 
  {
    alert("Please input number.");
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
      <td colspan="2" class="BlueClr">Leaflet Range</td>
    </tr>
    <tr> 
      <td colspan="2"></td>
    </tr>

    <tr> 
      <td colspan="2" bordercolor="#999966" class="ValidHead"></td>
    </tr>
<tr><td width="30%">First batch: </td>
      <td></td>
   </tr>

 <tr><td width="30%">Face Value: </td>
      <td>
<input name="Face_Value1" type=text value="<% = Face_Value1 %>" maxlength="3">	    
</td>
   </tr>
 <tr>
<td>Type:</td>
<td>
<input name="Coupon_Type1" type=text value="<% = Coupon_Type1 %>" maxlength="2">	    
</td>
</tr>
     
 <tr><td>Batch: </td>
      <td>
<input name="Batch1" type=text value="<% = Batch1 %>" maxlength="3">	    
</td>
   </tr>

<tr><td>Start Range:</td>
      <td>
<input name="Start_Range1" type=text value="<% = Start_Range1 %>" maxlength="6">	    
</td>
   </tr>
<tr><td>Number of Coupon:</td>
      <td>
<input name="NoOfCoupon1" type=text value="<% = NoOfCoupon1 %>" maxlength="2">	    
</td>
   </tr>

 <tr> 
      <td colspan="2" bordercolor="#999966" class="ValidHead"></td>
    </tr>
<tr><td width="30%">Second batch: </td>
      <td></td>
   </tr> 

 <tr><td width="30%">Face Value: </td>
      <td>
<input name="Face_Value2" type=text value="<% = Face_Value2 %>" maxlength="3">	    
</td>
   </tr>
 <tr>
<td>Type:</td>
<td>
<input name="Coupon_Type2" type=text value="<% = Coupon_Type2 %>" maxlength="2">	    
</td>
</tr>
     
 <tr><td>Batch: </td>
      <td>
<input name="Batch2" type=text value="<% = Batch2 %>" maxlength="3">	    
</td>
   </tr>

<tr><td>Start Range:</td>
      <td>
<input name="Start_Range2" type=text value="<% = Start_Range2 %>" maxlength="6">	    
</td>
   </tr>

<tr><td>Number of coupon:</td>
      <td>
<input name="NoOfCoupon2" type=text value="<% = NoOfCoupon2 %>" maxlength="2">	    
</td>
   </tr>

<tr> 
      <td colspan="2" bordercolor="#999966" class="ValidHead"></td>
    </tr>

<tr><td>Number of Leaflet:</td>
      <td>
<input name="NoOfLeaflet" type=text value="<% = NoOfLeaflet %>" maxlength="4">	    
</td>
   </tr>

	<tr>
<td></td>
<td>
  
</td>
	</tr>
     <tr>
 
    
                                        <td colspan="2" align="center"> 
<%  If RangeID <> "" Then %>
<input type="button" value="  Edit  " onClick="doEdit();" class="common">
<input type=hidden name=whatdo value="">
<% Else %>
<input type="button" value="  Submit  " onClick="dosubmit();" class="common">
<input type=hidden name=whatdo value="">
<% End If %>
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