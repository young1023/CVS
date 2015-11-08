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

    Separator = ""
    for i = 0 to RS.Fields.Count - 1
        Field = RS.Fields(i).Name
        if RX.Test(Field) then
            '
            ' According to recommendations:
            ' - Fields that contain CR/LF, Comma or Double-quote should be enclosed in double-quotes
            ' - Double-quote itself must be escaped by preceeding with another double-quote
            '
            Field = """" & Replace(Field, """", """""") & """"
        end if
        Response.Write Separator & Field
        Separator = ","
    next
    Response.Write vbNewLine

    '
    ' Writing the data rows
    '

    do until RS.EOF
        Separator = ""
        for i = 0 to RS.Fields.Count - 1
            '
            ' Note the concatenation with empty string below
            ' This assures that NULL values are converted to empty string
            '
            Field = RS.Fields(i).Value & ""
            if RX.Test(Field) then
                Field = """" & Replace(Field, """", """""") & """"
            end if
            Response.Write Separator & Field
            Separator = ","
        next
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

  


