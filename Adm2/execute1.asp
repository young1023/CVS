<!--#include file="include/SQLConn.inc" -->
<% 
'------------------------------------------------------------------------------------
'
'
' The page contain most of the create, modify and delete function 
'
'
'------------------------------------------------------------------------------------

response.expires=0
flag = trim(request.form("whatdo"))
'Response.write flag

''-----------------------------------------------------------------------------------------------------
'
'
'  Create new Pro coupon
'
'
'----------------------------------------------------------------------------------------------------- 


if flag = "add_pco" then
 
  Face_Value = int(replace(trim(request.form("Face_Value")),"'","''"))
  Product_Type = replace(trim(request.form("Product_Type")),"'","''")
  Batch = replace(trim(request.form("Batch")),"'","''")
  Start_Range = replace(trim(request.form("Start_Range")),"'","''")
  End_Range = request.form("End_Range")

  Start_Range1 = int(Start_Range)
  End_Range1   = int(End_Range)
  Digital = trim(request.form("Digital")) 
  Category = trim(request.form("Category"))
  Dealer = trim(request.form("Dealer")) 
  Canopy_Disc = trim(request.form("Canopy_Disc")) 
  Expiry_Date = trim(request.form("Expiry_Date")) 
  Issue_Date = trim(request.form("Issue_Date"))
  Completed = trim(request.form("Completed")) 
  Excel_Type = trim(request.form("Excel_Type")) 
  Effective_Date = trim(request.form("Effective_Date")) 
 
  sql1 = "Select count(1) as Tcount From CouponRequest Where Cast(FaceValue as float) = "& Face_Value

  sql1 = sql1 & " and Cast(Product_Type as float) = " & Product_Type

  sql1 = sql1 & " and Cast(batch as float) = " & Batch

  sql1 = sql1 & " and ( " & Start_Range1 & " between Cast(Start_Range as float) "

  sql1 = sql1 & " and Cast(End_Range as float)  " 

  sql1 = sql1 & " or  " & End_Range1 & " between Cast(Start_Range as float) "

  sql1 = sql1 & " and Cast(End_Range as float) "

  sql1 = sql1 & " or Cast(Start_Range as float) between " & Start_Range1 & " and " & End_Range1 

  sql1 = sql1 & " or Cast(End_Range as float)  between " & Start_Range1 & " and " & End_Range1 

  sql1 = sql1 & " ) " 

  'response.write sql1
  'response.end

  set rs1 = conn.execute(sql1)

  If rs1("Tcount") = 0 then

      sql = "Insert into CouponRequest (FaceValue, Product_Type, Batch, Start_Range, End_Range, Category, Dealer, Canopy_Copy_Disc, Expiry_Date, Issue_Date, Completed, Excel_Type, Effective_Date, Digital)"
      sql = sql & " values('"&Face_Value&"', '"&Product_Type&"', '"&Batch&"', '"&Start_Range&"', '"&End_Range&"', '"&Category&"', '"&Dealer&"', '"&Canopy_Disc&"', '"&Expiry_Date&"' , '"&Issue_Date&"', '"&Completed&"' , '"&Excel_Type&"', '"&Effective_Date&"', '"&Digital&"')"
  'response.write sql
  'response.end
     conn.execute(sql)
     message="The items are added."

  else

     message = "The range is covered by another pro coupon."

  end if
      whatgo = "pco1.asp"


'-----------------------------------------------------------------------------------------------------
'
'
'  Create new Pro coupon
'
'
'----------------------------------------------------------------------------------------------------- 


elseif flag = "edit_pco" then


  RequestID = Request("ID")
  Face_Value = replace(trim(request.form("Face_Value")),"'","''")
  Product_Type = replace(trim(request.form("Product_Type")),"'","''")
  Batch = replace(trim(request.form("Batch")),"'","''")
  Start_Range = replace(trim(request.form("Start_Range")),"'","''")
  End_Range = request.form("End_Range")
  Digital = request.form("Digital")
  Category = trim(request.form("Category"))
  Dealer = trim(request.form("Dealer")) 
  Canopy_Copy_Disc = trim(request.form("Canopy_Copy_Disc")) 
  Expiry_Date = trim(request.form("Expiry_Date")) 
  Issue_Date = trim(request.form("Issue_Date"))
  Completed = trim(request.form("Completed")) 
  Excel_Type = trim(request.form("Excel_Type")) 
  Effective_Date = trim(request.form("Effective_Date")) 
  
 
  sql = "Update CouponRequest Set FaceValue = '"&Face_Value&"', "

  sql = sql & "Product_Type = '"&Product_Type&"' ,"

  sql = sql & "Batch = '"&Batch&"', "

  sql = sql & "Start_Range = '"&Start_Range&"', End_Range = '"&End_Range&"', "

  sql = sql & "Digital = '"&Digital&"', "

  sql = sql & "Category = '"&Category&"', Dealer = '"&Dealer&"', "

  sql = sql & "Canopy_Copy_Disc = '"&Canopy_Copy_Disc&"', "

  sql = sql & "Expiry_Date = '"&Expiry_Date&"', Issue_Date = '"&Issue_Date&"' , "

  sql = sql & "Completed = '"&Completed&"', Excel_Type = '"&Excel_Type&"' ,"

  sql = sql & "Effective_Date = '"&Effective_Date&"' where RequestID ="&RequestID
      
  'response.write sql
  'response.end
     conn.execute(sql)
     message="Pro coupon is edited."
      whatgo = "pco2.asp?id="&RequestID



'==========================================================================
'
'
'  Delete Pro Coupon
'
'
'==========================================================================


elseif flag = "delpc" then
 
  delid=split(trim(request.form("mid")),",")
   for i=0 to Ubound(delid)
     sql="delete from CouponRequest where Requestid="&trim(delid(i))
     conn.execute(sql)
	 'response.write sql&"<br>"
   next
  message="Pro coupon record was deleted successfully"
   whatgo="pco_d1.asp"

'-----------------------------------------------------------------------------------------------------
'
'
'  new station
'
'
'----------------------------------------------------------------------------------------------------- 


elseif flag ="add_st" then

  
  station = replace(trim(request.form("station")),"'","''")
  ipaddress = replace(trim(request.form("ipaddress")),"'","''")
  stationname = replace(trim(request.form("StationName")),"'","''")
  MachineNo = replace(trim(request.form("MachineNo")),"'","''")
  SAPCode = replace(trim(request.form("SAPCode")),"'","''")
  ShipToCode = replace(trim(request.form("ShipToCode")),"'","''")
  SoldToCode = request.form("SoldToCode")
  Outdoor    = request.form("Outdoor")

      sql1 = "Select count(1) as Tcount From station where IPAddress ='" & ipaddress &"'"

      set rs1 = conn.execute(sql1)

      if rs1("Tcount") = 0 then


      sql = "Insert into station (station, IPAddress, StationName, MachineNo, SAPCode, SoldToCode, ShipToCode, Outdoor)"
      sql = sql & " values('"&station&"', '"&IPAddress&"', '"&StationName&"' , '"&MachineNo&"' , '"&SAPCode&"' , '"&SoldToCode&"', '"&ShipToCode&"', "&Outdoor&")"
  'response.write sql
  'response.end
     conn.execute(sql)
     message="Station is added."

     else

      message = "IP address is used."

    end if

 

      whatgo = "station1.asp"


'-----------------------------------------------------------------------------------------------------
'
'
'  edit station
'
'
'----------------------------------------------------------------------------------------------------- 


elseif flag ="edit_st" then

  
  station = replace(trim(request.form("station")),"'","''")
  ipaddress = replace(trim(request.form("ipaddress")),"'","''")
  stationname = replace(trim(request.form("StationName")),"'","''")
  MachineNo = replace(trim(request.form("MachineNo")),"'","''")
  SAPCode = replace(trim(request.form("SAPCode")),"'","''")
  ShipToCode = replace(trim(request.form("ShipToCode")),"'","''")
  SoldToCode = request.form("SoldToCode")
  Outdoor    = request.form("Outdoor")

      sql = "Update station set station='"&station&"', "

      sql = sql & "IPAddress='"&IPAddress&"', SAPCode='"&SAPCode&"' , "

      sql = sql & "StationName ='"&stationname&"', MachineNo = '"&MachineNo&"', "

      sql = sql & "SoldToCode='"&SoldToCode&"',  ShipToCode='"&ShipToCode&"', "

      sql = sql & "Outdoor="&Outdoor&" where IPAddress ='"& IPAddress &"'"

      'response.write sql
      'response.end
     conn.execute(sql)
     message="Station is edited."

 

      whatgo = "station1.asp"
'==========================================================================
'
'
'  Delete station
'
'
'==========================================================================


elseif flag = "del_st" then
 
  delid=split(trim(request.form("mid")),",")
   for i=0 to Ubound(delid)
     sql="delete from station where IPAddress='"&trim(delid(i))&"'"
     conn.execute(sql)
	 'response.write sql&"<br>"
   next
  message="Station was deleted successfully"
   whatgo="station1.asp"





'-----------------------------------------------------------------------------------------------------
'
'
'  new price
'
'
'----------------------------------------------------------------------------------------------------- 


elseif flag = "add_pr" then
 
  product = replace(trim(request.form("product")),"'","''")
  price = replace(trim(request.form("price")),"'","''")
  effectivedate = replace(trim(request.form("effectivedate")),"'","''")
  expirydate = replace(trim(request.form("expirydate")),"'","''")

      sql = "Insert into pricefile (product, unitprice, EffectiveDate, ExpiryDate)"
      sql = sql & " values('"&product&"', '"&price&"', '"&effectivedate&"', '"&ExpiryDate&"')"
  response.write sql
  'response.end
     conn.execute(sql)
     message="The items are added."
      whatgo = "price1.asp"


'==========================================================================
'
'
'  Delete price
'
'
'==========================================================================


elseif flag = "del_mtr" then
 
  delid=split(trim(request.form("mid")),",")
   for i=0 to Ubound(delid)
     sql="delete from mastercoupon where id='"&trim(delid(i))&"'"
     conn.execute(sql)
	 'response.write sql&"<br>"
   next
  message="Coupon was deleted successfully"
   whatgo="master1.asp"



end if


conn.close
set conn=nothing
wait=10

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href="include/hse.css" />
<meta http-equiv="Refresh" content="4; url='<%=whatgo%>'">
<title></title>
</head>
<body topmargin="0" marginwidth="0" marginheight="0" leftmargin="0" >
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

<p><a href='<%=whatgo%>'>Return</a></p>

<p><% response.write message %></p>

</td>
              </tr>
                </table>
                </td>
                </tr>
              </table>
           
              </body>

              </html>