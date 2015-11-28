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
 
  Face_Value = replace(trim(request.form("Face_Value")),"'","''")
  Product_Type = replace(trim(request.form("Product_Type")),"'","''")
  Batch = replace(trim(request.form("Batch")),"'","''")
  Start_Range = replace(trim(request.form("Start_Range")),"'","''")
  End_Range = request.form("End_Range")
  Category = trim(request.form("Category"))
  Dealer = trim(request.form("Dealer")) 
  Canopy_Disc = trim(request.form("Canopy_Disc")) 
  Expiry_Date = trim(request.form("Expiry_Date")) 
  Issue_Date = trim(request.form("Issue_Date"))
  Completed = trim(request.form("Completed")) 
  Excel_Type = trim(request.form("Excel_Type")) 
  Effective_Date = trim(request.form("Effective_Date")) 
  
 'response.write product_type
'response.end

      sql = "Insert into CouponRequest (FaceValue, Product_Type, Batch, Start_Range, End_Range, Category, Dealer, Canopy_Copy_Disc, Expiry_Date, Issue_Date, Completed, Excel_Type, Effective_Date)"
      sql = sql & " values('"&Face_Value&"', '"&Product_Type&"', '"&Batch&"', '"&Start_Range&"', '"&End_Range&"', '"&Category&"', '"&Dealer&"', '"&Canopy_Disc&"', '"&Expiry_Date&"' , '"&Issue_Date&"', '"&Completed&"' , '"&Excel_Type&"', '"&Effective_Date&"')"
  response.write sql
  'response.end
     conn.execute(sql)
     message="The items are added."
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

  sql = sql & "Category = '"&Category&"', Dealer = '"&Dealer&"', "

  sql = sql & "Canopy_Copy_Disc = '"&Canopy_Copy_Disc&"', "

  sql = sql & "Expiry_Date = '"&Expiry_Date&"', Issue_Date = '"&Issue_Date&"' , "

  sql = sql & "Completed = '"&Completed&"', Excel_Type = '"&Excel_Type&"' ,"

  sql = sql & "Effective_Date = '"&Effective_Date&"' where RequestID ="&RequestID
      
  'response.write sql
  'response.end
     conn.execute(sql)
     message="The items are edited."
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
  SAPCode = replace(trim(request.form("SAPCode")),"'","''")
  ShipToCode = replace(trim(request.form("ShipToCode")),"'","''")
  SoldToCode = request.form("SoldToCode")
  Outdoor    = request.form("Outdoor")

      sql = "Insert into station (station, IPAddress, SAPCode, SoldToCode, ShipToCode, Outdoor)"
      sql = sql & " values('"&station&"', '"&IPAddress&"', '"&SAPCode&"' , '"&SoldToCode&"', '"&ShipToCode&"', "&Outdoor&")"
  response.write sql
  'response.end
     conn.execute(sql)
     message="The items are added."

 

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
  SAPCode = replace(trim(request.form("SAPCode")),"'","''")
  ShipToCode = replace(trim(request.form("ShipToCode")),"'","''")
  SoldToCode = request.form("SoldToCode")
  Outdoor    = request.form("Outdoor")

      sql = "Update station set station='"&station&"', "

      sql = sql & "IPAddress='"&IPAddress&"', SAPCode='"&SAPCode&"' , "

      sql = sql & "SoldToCode='"&SoldToCode&"',  ShipToCode='"&ShipToCode&"', "

      sql = sql & "Outdoor="&Outdoor&" where IPAddress ='"& IPAddress &"'"

      'response.write sql
      'response.end
     conn.execute(sql)
     message="The items are edited."

 

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


elseif flag = "del_pr" then
 
  delid=split(trim(request.form("mid")),",")
   for i=0 to Ubound(delid)
     sql="delete from pricefile where product='"&trim(delid(i))&"'"
     conn.execute(sql)
	 'response.write sql&"<br>"
   next
  message="Price was deleted successfully"
   whatgo="price1.asp"



end if


conn.close
set conn=nothing
wait=10

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Refresh" content="3; url='<%=whatgo%>'">
<title></title>
</head>
<body topmargin="0" marginwidth="0" marginheight="0" leftmargin="0" >
 <br><br>
<table border=0 cellpadding=3 cellspacing=0 class=hardcolor width="90%" align=center>
  <tbody>
  <tr> 
    <td align=center bgcolor="#006699" height="28"><font color=white><%=message%></font></td>
  </tr>
  <tr>
    <td align=center height="38"><br><a href='<%=whatgo%>'>Return</a></td>
  </tr>
  </tbody>
</table>
</body>
</html>