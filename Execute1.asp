 <!--#include file="include/SQLConn.inc" -->
<%


   Set Rs1 = Server.CreateObject("Adodb.Recordset")
   Set Rs2 = Server.CreateObject("Adodb.Recordset")
   Set Rs3 = Server.CreateObject("Adodb.Recordset")
 
   StationID = Request("StationID")

   Barcode    = Request("Barcode")

   ProductType = Request("ProductType")

   CarPlate   = Request("CarPlate")

  

   response.write "Station ID:  " & Stationid & "<br>"

   response.write "Barcode: " & Barcode & "<br>"

   'response.end

   Set Message = Nothing

    Rs1.open ("Exec Checkrange '"&Barcode&"'") ,  conn,3,1

    Response.write ("Exec Checkrange '"&Barcode&"'")

   'Response.end

   If Rs1.EoF Then

           Message = "���ҥ��� - §��L�� "

 
    Response.Redirect  "CouponVerification2.asp?Message="&Message

 
   Else


       If Rs1("Expiry_Date") < DateValue(Now) Then

  
          Message = "���ҥ��� - §��L��!"


   Response.Redirect  "CouponVerification2.asp?Message="&Message


       End if

   End if


          Rs1.close
          set Rs1 = nothing

          Rs2.open ("Exec CheckCouponExist '"&Barcode&"'") ,  conn,3,1

          Response.write ("Exec CheckCouponExist '"&Barcode&"'")

          'iCount = Rs2("iCount")
         

          Rs2.close
          set Rs2 = nothing


          'response.end

        If Not Rs2.EoF Then

         Present_Date = Rs2("Present_Date")

         Message = "���ҥ��� - ���ƨϥ�!<br/><br/>"

       
         Else

        
          Message = "���Ҧ��\!"

     

          Rs3.Open ("Exec InsertCoupon '"&Barcode&"', '"&StationID&"','"&ProductType&"','"&CarPlate&"'"), Conn, 3, 1


          End if


         Response.Redirect  "CouponVerification2.asp?Message="&Message&"&CarPlate="&CarPlate&"&ProductType="&ProductType
       
%>
     