'******************************************************************
'Name: Pradip Vaghasiya - v1.4
'Script to get the Actual Working Hours for Day and Week wise
'To Run the script: Keep all files in Same Directory.
'Change Employee ID Wherever Required
'******************************************************************
'******************************************************************
'Name: Asawari Buche - v2.6
'abhishekga@cybage.com
'Script to get the Actual Working Hours for Day, Week wise and month
'To Run the script: Keep all files in Same Directory.
'Change Employee ID below
'******************************************************************

employeeID = "15282"
numberOfResultColumn = 5
columnNoForMachine = 2
columnNoForDirection = 3
columnNoForTime = 4
LakshmanRekhaArray = Array("Tripod", "Flap Barrier", "Basement Entrance")

set IE = CreateObject("InternetExplorer.Application") 
IE.Visible = True

path = "http://cybmis-app-3-ld/Report Builder/RPTN/Reportpage.aspx"
'IE.Navigate "http://cybagemis.cybage.com/Report Builder/RPTN/ReportPage.aspx"
IE.Navigate path
waitToLoad IE

GoToLink("Today's and Yesterday's Swipe Log")
WScript.sleep 300
waitToLoad IE

IE.Document.GetElementById("DayDropDownList8665").selectedIndex = 1
AWHSecondsToday = calculateAWH()
AWHRemainingSecs = 28800 - AWHSecondsToday

If (AWHRemainingSecs) > 0 Then
	'TodaysSwipesText = "    TODAY" & vbCrLf & "  -------------------------------------" & vbCrLf & "  Working Hours :-  " & showAWH(AWHSecondsToday) & vbCrLf & _
	TodaysSwipesText = "TODAY" & vbCrLf & "  Working Hours :-  " & showAWH(AWHSecondsToday) & vbCrLf & _
	"  Probable End Time (8 Hours) :-  " & DateAdd("s", AWHRemainingSecs, Now()) & vbCrLf & _
	"  Remaining Hours :-  " & showAWH(AWHRemainingSecs) & vbCrLf & vbCrLf
Else
	'TodaysSwipesText = "    TODAY" & vbCrLf & "  -------------------------------------" & vbCrLf & "  Working Hours :-  " & showAWH(AWHSecondsToday) & vbCrLf & vbCrLf
	TodaysSwipesText = "    TODAY" & vbCrLf & "  Working Hours :-  " & showAWH(AWHSecondsToday) & vbCrLf & vbCrLf
End If

'IE.Navigate "http://cybagemis.cybage.com/Report Builder/RPTN/ReportPage.aspx"
IE.Navigate path
WScript.sleep 300
waitToLoad IE

GoToLink("Today's and Yesterday's Swipe Log")
WScript.sleep 300
waitToLoad IE

On Error Resume Next

IE.Document.GetElementById("DayDropDownList8665").selectedIndex = 2
AWHSecondsYesterday = calculateAWH()

If ((Weekday(Now()) > 2) and (Weekday(Now()) < 7)) and (AWHSecondsYesterday = 0) then
	AWHSecondsYesterday = 28800
end if

'YesterdaysSwipesText = "    YESTERDAY" & vbCrLf & "  -------------------------------------" & vbCrLf & "  Working Hours :-  " & showAWH(AWHSecondsYesterday) & vbCrLf & vbCrLf
YesterdaysSwipesText = "YESTERDAY Working Hours :-  " & showAWH(AWHSecondsYesterday) & vbCrLf & vbCrLf

'MonthlyHourText = "MONTHLY - " & MonthName(Month(Now())) & " " & Year(Now()) & vbCrLf & "  -------------------------------------" & vbCrLf & calculateMONTHLY()
MonthlyHourText = "MONTHLY - " & MonthName(Month(Now())) & " " & Year(Now()) & vbCrLf & calculateMONTHLY()
TotalAWH = TodaysSwipesText & MonthlyHourText
IE.quit
MsgBox(TotalAWH)

Function calculateMONTHLY
	numberOfResultColumn = 11
	columnNoForDate = 2
	columnNoForAWH = 7
	
	'IE.Navigate "http://cybagemis.cybage.com/Core/Common/Introduction.aspx"
	IE.Navigate path
	WScript.sleep 300
	waitToLoad IE
	
	StartDate = "01-" & MonthName(DatePart("m", now), True) & "-" & Year(now)
	'dt = DateAdd("d", -2, date)
	'EndDate = Right("0" & Day(dt), 2) & "-" & MonthName(Month(dt), true) & "-" & Year(dt)
	
	'IE.Navigate "http://cybagemis.cybage.com/Report Builder/RPTN/ReportPage.aspx"
	IE.Navigate path
	WScript.sleep 300
	waitToLoad IE
	
	GoToLink("Attendance Log Report")
	WScript.sleep 300
	waitToLoad IE
	WScript.sleep 500
	IE.Document.GetElementById("DMNDateDateRangeControl4392_FromDateCalender").value = startDate
	IE.Document.GetElementById("DMNDateDateRangeControl4392_FromDateCalender_DTB").value = startDate
	IE.Document.GetElementById("DMNDateDateRangeControl4392_FromDateCalender_hf").value = startDate
	'IE.Document.GetElementById("DMNDateDateRangeControl4392_ToDateCalender").value = EndDate
	'IE.Document.GetElementById("DMNDateDateRangeControl4392_ToDateCalender_DTB").value = EndDate
	'IE.Document.GetElementById("DMNDateDateRangeControl4392_ToDateCalender_hf").value = EndDate
	waitToLoad IE
	
	IE.Document.GetElementById("ViewReportImageButton").click
	waitToLoad IE
	
	ReDim inOutTimings(numberOfResultColumn, 1)
	resultColumn = 0
	result = -1
	
	startShowing = False
	For page = 1 to IE.Document.GetElementById("ReportViewer1").ClientController.TotalPages
		resultColumn = 0
		result = result + 1
		redim preserve inOutTimings(numberOfResultColumn, result + 1)
		
		For each tableTag in IE.Document.GetElementsByTagName("TABLE")
			For each row in tableTag.Rows
				For each cell in row.cells
					if cell.innerText = employeeID then
						startShowing = True
					end if
					
					if startShowing Then
						if resultColumn = numberOfResultColumn then
							if cell.innerText = employeeID then
								resultColumn = 0
								result = result + 1
								redim preserve inOutTimings(numberOfResultColumn, result + 1)
							else
								startShowing = False
								Exit For
							end if
						end if
						
						inOutTimings(resultColumn, result) = cell.innerText
						resultColumn = resultColumn + 1
					end if
				Next
			Next
		Next
		
		IE.Document.GetElementById("ReportViewer1_ctl01_ctl01_ctl05_ctl00").click
		waitToLoad IE
	Next
	
	'IE.quit
	
	Total = 0
	hrs = 0
	hlfhrs = 0
	
	i_all = 0
	i_hrs = 0
	i_min = 0
	
	i_mnthmins = 0
	i_mnthhrs = 0
	i_mnthTot = 0
	
	workingDaysCount = 0
	'result = result - 1		'For handling yesterday - calculation done for this separately

	For row = 0 To result
		currentAWH = inOutTimings(columnNoForAWH, row)
		currentDate = CDate(inOutTimings(columnNoForDate, row))
		currentDay = Day(currentDate)
		if (currentDay < 10) then 
			currentDay = "0" & currentDay
		end if
		
		'Saturday - calculate weeks total and display it
		if (Weekday(currentDate) = 7) then
			i_hrs = i_hrs + Int(i_min/ 60)
			i_min = i_min mod 60
			
			Total = i_hrs & ":" & Right("0" & i_min, 2)
			TotalAWH = TotalAWH & " " & currentDay & " - Sat  -  Weeks Total:- " & Total & vbCrLf
			
			'Reset weekly variables
			Total = 0
			i_hrs = 0
			i_min = 0
		
		'Weekdays
		elseif ((Weekday(currentDate) > 1) and (Weekday(currentDate) < 7)) then
			workingDaysCount = workingDaysCount + 1
			hrs = Replace(currentAWH, ":", ".")
			statusAWH = "	- "
			hrsStr = 0
			minStr = 0
			
			'check for 1st half leave/holiday
			statusFirstHalf = inOutTimings(columnNoForAWH + 1, row)
			if (statusFirstHalf <> " " and (InStr(statusFirstHalf, "Leave") or InStr(statusFirstHalf, "Holiday") or InStr(statusFirstHalf, "LWP"))) then
				hrsStr = hrsStr + 4
				minStr = minStr + 0
				statusAWH = statusAWH & "1st half " & statusFirstHalf
			end if
			
			'check for 2nd half leave/holiday
			statusSecondHalf = inOutTimings(columnNoForAWH + 2, row)
			if (statusSecondHalf <> " " and (InStr(statusSecondHalf, "Leave") or InStr(statusSecondHalf, "Holiday") or InStr(statusSecondHalf, "LWP"))) then
				hrsStr = hrsStr + 4
				minStr = minStr + 0
				statusAWH = statusAWH & "2nd half " & statusSecondHalf
			end if
			
			'Separate hours and minutes Current AWH
			if (hrs <> " ") then
				i_all = Replace(currentAWH, ":", "")
				hrsStr = hrsStr + Int(i_all/ 100)
				minStr = minStr + i_all mod 100
			elseif (row = result) then
				hrsStr = hrsStr + Int(AWHSecondsYesterday/3600)
				minStr = minStr + Abs((AWHSecondsYesterday/60) mod 60)
				statusAWH = statusAWH & "YESTERDAY"
			end if
			
			i_hrs = i_hrs + hrsStr				'Add hour to Weekly
			i_min = i_min + minStr				'Add minutes to Weekly
			i_mnthhrs = i_mnthhrs + hrsStr		'Add hour to Monthly
			i_mnthmins = i_mnthmins + minStr	'Add minutes to Monthly

			'Append 0 to hour if it is below 10
			if (hrsStr < 10) then
				hrsStr = "0" & hrsStr
			end if
			
			'Append 0 to minute if it is below 10
			if (minStr < 10) then
				minStr = "0" & minStr
			end if

			'If both half have same status show only one (Except for yesterday)
			if ( (StrComp(statusFirstHalf,statusSecondHalf) = 0) and (row < result) ) then
				statusAWH = "	- " & statusFirstHalf
			end if

			'Display hours for date along with status
			TotalAWH = TotalAWH & " " & currentDay & " - " & hrsStr & ":" & minStr & statusAWH & vbCrLf
		
		'Sunday
		else
			TotalAWH = TotalAWH & vbCrLf & " " & currentDay & " - Sun" & vbCrLf
		end if
		
	Next
	
	'Month's Total
	'AWHInSeconds = AWHSecondsToday + AWHSecondsYesterday
	AWHInSeconds = AWHSecondsToday
	Hours = Int(AWHInSeconds/3600)
	remainderSeconds = AWHInSeconds mod 3600
	Minutes = Int(remainderSeconds/60)
	
	i_mnthmins = i_mnthmins + Minutes
	i_mnthhrs = i_mnthhrs + Hours + Int(i_mnthmins/60)
	i_mnthmins = i_mnthmins mod 60
	
	i_mnthTot = i_mnthhrs & ":" & Right("0" & i_mnthmins, 2)
	
	REM if (Weekday(Now()) = 2) then
		REM workingDaysCount = workingDaysCount + 1
	REM elseif ((Weekday(Now()) > 2) and (Weekday(Now()) < 7)) then
		REM workingDaysCount = workingDaysCount + 2
	REM end if
	
	workingDaysCount = workingDaysCount + 1		'increment workingDaysCount to add today's day
	expectedHours = workingDaysCount * 8
	minutesDiff = (expectedHours * 60) - (i_mnthhrs * 60) - i_mnthmins
	hrsDiffArr = Split(CStr(FormatNumber(minutesDiff/60)), ".")
	minDiffArr = Right("0" & Abs(minutesDiff mod 60), 2)
	
	TotalAWH = TotalAWH & vbCrLf & "SUMMARY"
	TotalAWH = TotalAWH & vbCrLf & " Month's Total: " & i_mnthTot
	TotalAWH = TotalAWH & vbCrLf & " Expected Total till today: " & expectedHours & ":00"
	TotalAWH = TotalAWH & vbCrLf & " Remaining hours for Total: " & hrsDiffArr(0) & ":" & minDiffArr
	TotalAWH = TotalAWH & vbCrLf & " Adjusted end time: " & DateAdd("s", (minutesDiff * 60), Now())
	
	calculateMONTHLY = TotalAWH
End Function

Function calculateAWH
	IE.Document.GetElementById("ViewReportImageButton").click
	waitToLoad IE
	
	Dim inOutTimings()
	
	ReDim inOutTimings(numberOfResultColumn, 1)
	resultColumn = 0
	result = -1
	
	startShowing = False
	For page = 1 to IE.Document.GetElementById("ReportViewer1").ClientController.TotalPages
		resultColumn = 0
		result = result + 1
		redim preserve inOutTimings(numberOfResultColumn, result + 1)
		
		For each tableTag in IE.Document.GetElementsByTagName("TABLE")
			For each row in tableTag.Rows
				For each cell in row.cells
					if cell.innerText = employeeID then
						startShowing = True
					end if
					
					if startShowing Then
						if resultColumn = numberOfResultColumn then
							if cell.innerText = employeeID then
								resultColumn = 0
								result = result + 1
								redim preserve inOutTimings(numberOfResultColumn, result + 1)
							else
								startShowing = False
								Exit For
							end if
						end if
						
						inOutTimings(resultColumn, result) = cell.innerText
						resultColumn = resultColumn + 1
					end if
				Next
			Next
		Next
		
		IE.Document.GetElementById("ReportViewer1_ctl01_ctl01_ctl05_ctl00").click
		waitToLoad IE
	Next
	
	dim startDateTime, AWHInSeconds
	PersonInsideLakshmanRekha = False
	AWHInSeconds = 0
	PersonEntered = False
	PersonExited = True
	startDateTime = CDate(DateValue(Now()) & " 00:00:00")
	ExitDateTime = startDateTime
	DeadLine = CDate(DateValue(Now()) & " 09:00:00 PM")
	'MsgBox(deadline)
	
	For row = 0 To result
		For Each LakshmanRekhaArrayKeyWord in LakshmanRekhaArray
			if InStr(inOutTimings(columnNoForMachine, row), LakshmanRekhaArrayKeyWord) > 0 then
				if inOutTimings(columnNoForDirection, row) = "Entry" then
					If Not PersonEntered and PersonExited Then
						If DateDiff("s", DeadLine, startDateTime) > 0 Then
							startDateTime = DeadLine
							ExitDateTime = DeadLine
						ElseIf DateDiff("s", DeadLine, ExitDateTime) > 0 Then
							ExitDateTime = DeadLine
						End If
						AWHInSeconds = AWHInSeconds + DateDiff("s", startDateTime, ExitDateTime)
					End If
					
					If PersonExited Then
						startDateTime = CDate(DateValue(Now()) & " " & inOutTimings(columnNoForTime, row))
					End If
					
					PersonInsideLakshmanRekha = True
					PersonExited = False
					PersonEntered = True
				else
					ExitDateTime =  CDate(DateValue(Now()) & " " & inOutTimings(columnNoForTime, row))
					PersonInsideLakshmanRekha = False
					PersonExited = True
					PersonEntered = False
				end if
			end if
		Next
	Next
	
	REM if PersonInsideLakshmanRekha then
		REM AWHInSeconds = AWHInSeconds + DateDiff("s", startDateTime, Now())
	REM Else
		REM AWHInSeconds = AWHInSeconds + DateDiff("s", startDateTime, ExitDateTime)
	REM End If
	
	If PersonInsideLakshmanRekha Then
		ExitDateTime = Now()
	End If
	
	If DateDiff("s", DeadLine, startDateTime) > 0 Then
		startDateTime = DeadLine
		ExitDateTime = DeadLine
	ElseIf DateDiff("s", DeadLine, ExitDateTime) > 0 Then
		ExitDateTime = DeadLine
	End If
	
	AWHInSeconds = AWHInSeconds + DateDiff("s", startDateTime, ExitDateTime)
	
	calculateAWH = AWHInSeconds
	
End Function

Function showAWH(AWHInSeconds)
	AWHMins = Int(AWHInSeconds/60)
	
	Hours = Int(AWHMins/60)
	Minutes = Abs(AWHMins mod 60)
	Seconds = Abs(AWHInSeconds mod 60)
	
	showAWH = Hours & ":" &_
				Right("0" & Minutes, 2) & ":" &_
				Right("0" & Seconds, 2)
	
End Function

Sub GoToLink(innerText)
	For each aTag in IE.Document.getElementsByTagName("a")
		If aTag.innerText = innerText Then
			aTag.click
		End If
	Next
End Sub

Function waitToLoad(IE)
    maxTry = 50
    try = 1
    WScript.sleep 300
    Do While (try < maxTry AND IE.busy)
    	WScript.sleep 100
     	try = try + 1
    Loop
	
End Function