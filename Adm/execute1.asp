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
flage=trim(request.form("whatdo"))
'Response.write flage

''-----------------------------------------------------------------------------------------------------
'
'
'  Create new Pro coupon
'
'
'----------------------------------------------------------------------------------------------------- 


if flage="add_pco" then
 
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
  
 response.write product_type
'response.end

      sql = "Insert into CouponRequest (FaceValue, Product_Type, Batch, Start_Range, End_Range, Category, Dealer, Canopy_Copy_Disc, Expiry_Date, Issue_Date, Completed, Excel_Type, Effective_Date)"
      sql = sql & " values('"&Face_Value&"', '"&Product_Type&"', '"&Batch&"', '"&Start_Range&"', '"&End_Range&"', '"&Category&"', '"&Dealer&"', '"&Canopy_Copy_Disc&"', '"&Expiry_Date&"' , '"&Issue_Date&"', '"&Completed&"' , '"&Excel_Type&"', '"&Effective_Date&"')"
  response.write sql
  'response.end
     conn.execute(sql)
     message="The items are added."
      whatgo = "pco1.asp"


'==========================================================================
'
'
'  Delete Pro Coupon
'
'
'==========================================================================


elseif flage="delpc" then
 
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


elseif flage="add_st" then

  station = replace(trim(request.form("station")),"'","''")

  Sql1 = "Select station from station where station='"&station&"'"

   'response.write sql1

  Set Rs1 = conn.execute(sql1)

  If Rs1.EoF Then
  
  ipaddress = replace(trim(request.form("ipaddress")),"'","''")
  SAPCode = replace(trim(request.form("SAPCode")),"'","''")
  ShipToCode = replace(trim(request.form("ShipToCode")),"'","''")
  SoldToCode = request.form("SoldToCode")

      sql = "Insert into station (station, IPAddress, SAPCode, SoldToCode, ShipToCode)"
      sql = sql & " values('"&station&"', '"&IPAddress&"', '"&SAPCode&"' , '"&SoldToCode&"', '"&ShipToCode&"')"
  response.write sql
  'response.end
     conn.execute(sql)
     message="The items are added."

   Else

     message="Station number is used."

   End if

      whatgo = "station.asp"

'==========================================================================
'
'
'  Delete station
'
'
'==========================================================================


elseif flage="del_st" then
 
  delid=split(trim(request.form("mid")),",")
   for i=0 to Ubound(delid)
     sql="delete from station where station='"&trim(delid(i))&"'"
     conn.execute(sql)
	 'response.write sql&"<br>"
   next
  message="Station was deleted successfully"
   whatgo="station.asp"





'-----------------------------------------------------------------------------------------------------
'
'
'  new price
'
'
'----------------------------------------------------------------------------------------------------- 


elseif flage="add_pr" then
 
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


elseif flage="del_pr" then
 
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