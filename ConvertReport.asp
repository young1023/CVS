<!--#include file="include/SQLConn.inc" -->
<%
sub Write_CSV_From_Recordset(RS)

   
    if RS.EOF then
    
        '
        ' There is no data to be written
        '
        exit sub
    
    end if

    dim RX
    set RX = new RegExp
        RX.Pattern = "\r|\n|,|"""

    dim i
    dim Field
    dim Separator

    '
    ' Writing the header row (header row contains field names)
    '

    Separator = ","
   
        Response.Write "Present Date and Time"
        Response.Write Separator 
        Response.Write "Period"
        Response.Write Separator 
        Response.Write "Number"
        Response.Write Separator
        Response.Write "Type"
        Response.Write Separator
        Response.Write "Product"
        Response.Write Separator
        Response.Write "Litre"
        Response.Write Separator
        Response.Write "Amount"
        Response.Write Separator
        Response.Write "Car"

    Response.Write vbNewLine

    '
    ' Writing the data rows
    '

    do until RS.EOF
     
        Response.Write Rs("Present_Date")
        Response.Write Separator 
        Response.Write Rs("Period")
        Response.Write Separator 
        Response.Write Rs("Coupon_Number")
        Response.Write Separator
        Response.Write Rs("Coupon_Type")
        Response.Write Separator
        Response.Write Rs("Product_Type")
        Response.Write Separator
        Response.Write Rs("SaleLitre")
        Response.Write Separator
        Response.Write RS("SaleAmount")
        Response.Write Separator
        Response.Write Rs("Car_ID")

        Response.Write vbNewLine
        RS.MoveNext
    loop

end sub

'
' EXAMPLE USAGE
'
' - Open a RECORDSET object (forward-only, read-only recommended)
' - Send appropriate response headers
' - Call the function
'

dim RS1


UserID        = trim(Request("UserID"))

Level         = trim(Request("SecLevel"))

StationID     = trim(Request("StationID"))

UserID        = Rtrim(Request("UserID"))

SDay          = Request("SDay")

SMonth        = Request("SMonth") 

SYear         = Request("SYear")

NDay          = Request("NDay")

NMonth        = Request("NMonth") 

NYear         = Request("NYear")


Search_Date = Formatdatetime(DateSerial(SYear, SMonth, SDay),2)

Search_NDate = Formatdatetime(DateSerial(NYear, NMonth, NDay),2)

      
      fsql = "select * from MasterCoupon where RequestedID = "&StationID &_


       " and  Present_Date >=   Convert(datetime, '" & Search_Date &"', 105) " &_

  
       " and  Present_Date < DATEADD(dd,DATEDIFF(dd,0, Convert(datetime, '" & Search_NDate &"', 105)),0) + 1 " &_

 
       " order by id desc"

        'response.write fsql
        'response.end

         Set Rs1 = Conn.Execute(fsql)

     Response.ContentType = "text/csv"

     Today = formatdatetime(now(),2)

     Response.AddHeader "Content-Disposition", "attachment;filename=Â§¨é¬ö¿ı("&Today&").csv"

     Write_CSV_From_Recordset RS1

		 
%>

  


