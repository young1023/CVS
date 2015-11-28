<%
'*************************************************************************************************
'author : robin coode
'email  : robin@dbnetlink.com
'www    : www.dbnetlink.co.uk
'*************************************************************************************************
'VBScript Calendar class.

'let properties
'		Mnth								Calendar month
'		Yr									Calendar year
'		FontSize					
'		Columns							No of colums for year calendar (1,2,3,4,6 or 12)
'		FontFace				
'		FontColour		
'		FillColour					Background colour for "S M T W T F S"
'		BorderColour	
'		BackgroundColour	
'		FullYearLink				ASP page for full year calender link
'
'get properties
'		MonthCal						HTML table string		
'		YearCal							HTML table string

'methods
'	LoadMonthArray				Private
'*************************************************************************************************
%>
<script>
function showyearcal(link, year) {
if (link.indexOf('?') > 0) 
	link = link + '&year=' + year
else
	link = link + '?year=' + year
calwin=window.open( link, 'calwin', 'toolbar=yes, scrollbars=yes, status=yes, width=680, height=480' )
if (typeof(calwin.focus) != "undefined") {
	calwin.focus()
	}
}
function changemonth(moveby) { 
document.calform.calmonth.value = document.calform.calmonth.value - 0 + moveby;
document.calform.submit(); 
}
function changeyear(moveby) { 
document.calform.calyear.value = document.calform.calyear.value - 0 + moveby; 
document.calform.submit(); 
}
</script>
<style> 
td.day {font-family:arial;font-size:8pt;color:black} 
</style>
<%
class calendar
	private M, Y, D, WeekNo, MonthArray, FSize, FFace, FColour, BorderCol, FillCol, BGCol, SingleMonth, FYLink, Cols, cStyleSheet

	property let Mnth(Month)  
		if Month >= 1 and Month <= 12 then
			M = Month 
		end if
	end property

	property let Yr(Year)	
		if Year > 1 and Year < 9999 then
			Y = Int(Year)
		end if
	end property
	
	property let FontSize(FS)
		if FS >= 1 and FS <= 7 then
			FSize = FS
		end if
	end property

	property let Columns(C)
		select case C
		case 1,2,3,4,6,12
			Cols = C
		case else
			Cols = 4
		end select
	end property

	property let FontFace(FF)	
		if FF <> "" then
			FFace = FF
		end if
	end property

	property let FontColour(FC)	
		if FC <> "" then
			FColour = FC 
		end if
	end property

	property let FillColour(FC)	
		if FC <> "" then
			FillCol = FC 
		end if
	end property

	property let BorderColour(BC)	
		if BC <> "" then
			BorderCol = BC 
		end if		
	end property

	property let BackgroundColour(BGC) 
		if BGC <> "" then
			BgCol = BGC 
		end if
	end property
	
	property let FullYearLink(FYL) FYLink = FYL end property
	property let StyleSheet(SS) cStyleSheet = SS end property

  private Sub Class_Initialize
    Mnth = Month(Now)
		Yr = Year(Now)
		FFace = "arial"
		FSize = 2
		FColour = "black"
		BorderCol = "steelblue"
		FillCol = "lightblue"
		BgCol = "gainsboro"
		SingleMonth = true
		FYLink = ""
		Cols = 4
		StyleSheet = false
  End Sub

  private Sub LoadMonthArray
		Dim Dte, FirstDayNo

		Redim MonthArray(6,7)

		for D = 1 to 31
			Dte = DateSerial(Y,M,D)
			if D = 1 then
				FirstDayNo = Weekday(Dte)
			end if
			if M = Month(Dte) and D = Day(Dte) then
				WeekNo = Abs( Int( ( ( FirstDayNo + D -1 ) /7 )*-1) )
				MonthArray( Weekno, Weekday(Dte) ) = D
			end if
		next
	end sub

	property get MonthCal
		dim HTML, FontStr, Colour, ColSpan

		if Request.Form("calmonth") <> "" then
			M = Int( Request.Form("calmonth") )
			Y = Int( Request.Form("calyear") )
			if M > 12 then
				M = 1
				Y = Y + 1
			end if
			if M < 1 then
				M = 12
				Y = Y -1
			end if
		end if

		LoadMonthArray

		FontStr = "<font face=""" & FFace & """ size=" & FSize & " color=" & FColour & ">"

		HTML = "<table cellspacing=	0 bgcolor=" & BgCol & " bordercolor=" & BorderCol & " border=1 width=""100%"">"

		HTML = HTML & "<tr>"		
		if SingleMonth then
			HTML = HTML & "<form name=calform method=post>"
			HTML = HTML & "<td align=center>" & FontStr & "<a href=javascript:changemonth(-1)>&lt;</a></td>"
			HTML = HTML & "<td align=center colspan=5>" & FontStr & MonthName(M)
			if FYLink <> "" then
				HTML = HTML & "&nbsp;<a href=javascript:showyearcal('" & Server.URLEncode(FYLink) & "',"& Y & ")>" & Y & "</a>"
			else
				HTML = HTML & " " & Y
			end if
			HTML = HTML & "</font></td>"
			HTML = HTML & "<td align=center>" & FontStr & "<a href=javascript:changemonth(1)>&gt;</a></td>"
		else
			HTML = HTML & "<td align=center colspan=7>" & FontStr & MonthName(M) & "</td>"
		end if

		HTML = HTML & "</tr>" 

		for D = 1 to 7
			HTML = HTML & "<th width=""14%"" bgcolor=" & FillCol & ">" & FontStr & Left(WeekdayName(d),1) & "</font></th>"
		next

		for WeekNo = 1 to 6
			HTML = HTML & "<tr>"
			for D = 1 to 7
				HTML = HTML & "<td "
				
				if cStyleSheet then
					HTML = HTML & "class=day "
				end if

				if MonthArray(WeekNo,D) = "" then
					MonthArray(WeekNo,D) = "&nbsp;"
				else
					if Date = DateSerial(Y,M,MonthArray(WeekNo,D)) then
						HTML = HTML & "bgcolor=" & BorderCol 
						FontStr = Replace( FontStr, FColour, BgCol )
					end if
				end if

				if not cStyleSheet then
					HTML = HTML & ">" & FontStr & MonthArray(WeekNo,D) & "</font></td>"
				else
					HTML = HTML & ">" & MonthArray(WeekNo,D) & "</td>"
				end if

				if IsNumeric( MonthArray(WeekNo,D) ) then
					if Date = DateSerial(Y,M,MonthArray(WeekNo,D)) then
						FontStr = Replace( FontStr, BgCol, FColour )
					end if
				end if
			next
			HTML = HTML & "</tr>"
		next

		if SingleMonth then
			HTML = HTML & "<input type=hidden name=calmonth value=" & M & "></input>"
			HTML = HTML & "<input type=hidden name=calyear value=" & Y & "></input>"

			HTML = HTML & "</form>"
		end if

		HTML = HTML & "</table>" 

		MonthCal = HTML

	end property

	property get YearCal
		Dim HTML, Col, Row, MonthSave, Rows

		MonthSave = M
		SingleMonth = false

		if Request.Form("calyear") <> "" then
			Yr = Request.Form("calyear")
		end if

		Rows = 12/Cols

		HTML = HTML & "<table border=0><form name=calform method=post>"
		HTML = HTML & "<tr><td align=center colspan=" & Cols & ">"
		HTML = HTML & "<font face=""" & FFace & """ size=6 color=" & FColour & ">"
		
		if not CStyleSheet then
			HTML = HTML & "<a href=javascript:changeyear(-1)>&lt;</a>&nbsp;" & Y & "&nbsp;<a href=javascript:changeyear(1)>&gt;</a>"
		else
			HTML = HTML & Y 
		end if
		
		HTML = HTML & "</font></td></tr>"

		for Row = 1 to Rows
			HTML = HTML & "<tr>"
			for Col = 1 to Cols
				Mnth = Col + ((Row -1) * Cols)
				HTML = HTML & "<td>" & MonthCal & "</td>"
			next
			HTML = HTML & "</tr>"
		next

		HTML = HTML & "<input type=hidden name=calyear value=" & Y & "></input></form></table>"

		Mnth = MonthSave

		YearCal = HTML
	end property

end class
%>