 <!--#include file="include/SQLConn.inc" -->
<%

   IPAddress = Request.ServerVariables("REMOTE_ADDR")

   ScanDate = Now()

   Set Rs0 = Server.CreateObject("Adodb.Recordset")
   Set Rs1 = Server.CreateObject("Adodb.Recordset")
   Set Rs2 = Server.CreateObject("Adodb.Recordset")
   Set Rs3 = Server.CreateObject("Adodb.Recordset")
 
   Barcode    = trim(Request("Barcode"))

   ProductType = Request("ProductType")
 

   'response.write "Station ID:  " & Stationid & "<br>"

   'response.write "Barcode: " & Barcode & "<br>"

   'response.end

   Set Message = Nothing

   ' Check Coupon Type
  
   Rs0.open ("Exec Check_CouponType '"&Barcode&"', '"&ProductType&"'") ,  conn,3,1

   'Response.write ("Exec Check_CouponType '"&Barcode&"', '"&ProductType&"'")

   'Response.end

   If Rs0("Tcount") = 0 Then

           Message = "���ҥ��� - ���~��������!"


    Response.Redirect  "CouponVerification.asp?ProductType="&ProductType&"&Color=2&Message="&Message
      
 
   End if


   ' Check Barcode Range

    Rs1.open ("Exec Checkrange '"&Barcode&"'") ,  conn,3,1

   Response.write ("Exec Checkrange '"&Barcode&"'")

   'Response.end

   If Rs1.EoF Then

           Message = "���ҥ��� - §��L�� "


    Response.Redirect  "CouponVerification.asp?ProductType="&ProductType&"&Color=2&Message="&Message
      
 
   Else


       If Rs1("Expiry_Date") < DateValue(Now) Then

  
          Message = "���ҥ��� - §��L��!"
        
          

    Response.Redirect  "CouponVerification.asp?ProductType="&ProductType&"&Color=2&Message="&Message
      

       End if

   End if


          Rs1.close
          set Rs1 = nothing

          Rs2.open ("Exec CheckCouponExist '"&Barcode&"'") ,  conn,3,1

          Response.write ("Exec CheckCouponExist '"&Barcode&"'")

     
         If Rs2.EoF Then


          Response.write ("Exec InsertCoupon '"&Barcode&"', '"&IPAddress&"','"&ProductType&"','"&ScanDate&"'")

          Rs3.Open ("Exec InsertCoupon2 '"&Barcode&"', '"&IPAddress&"','"&ProductType&"','"&ScanDate&"'"), Conn, 3, 1

               If Not Rs3.EoF Then              
          
               Message = "���Ҧ��\! ���X: " & Rs3("Coupon_number") & " �ɶ�: " & Rs3("Present_Date")
              

                Response.Redirect  "CouponVerification.asp?ProductType="&ProductType&"&Color=1&Message="&Message
    
               
                   Else


                Message = "System busy. Please try again. "
               

                Response.Redirect  "CouponVerification.asp?ProductType="&ProductType&"&Color=2&Message="&Message
    

               End


     
         Else

         Present_Date = Rs2("Present_Date")

         RequestedID  = Rs2("RequestedID")

         StationName  = Rs2("StationName")

         'response.write Present_Date
         'response.end
         Color   = "2"

         Message = "���ҥ��� - ���ƨϥ�! §�� " & Barcode & " ���� " & Present_Date & " �b"& StationName & "������!"


         Response.Redirect  "CouponVerification.asp?ProductType="&ProductType&"&Color="&Color&"&Message="&Message
      

       
          End if


          Rs2.close
          set Rs2 = nothing

 
%>
     