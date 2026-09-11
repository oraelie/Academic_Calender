Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Globalization
Imports ExcelDataReader

Public Class AcademicCalendarFeed
    Inherits System.Web.UI.Page

    Private ReadOnly Property ExcelFilePath As String
        Get
            Return Server.MapPath("~/App_Data/AcademicCalendar.xlsx")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

        Try

            Dim dt As DataTable = ReadExcelEvents()
            Dim icsContent As String = BuildICSContent(dt)
            Dim fileBytes As Byte() = Encoding.UTF8.GetBytes(icsContent)

            Dim excelLastModifiedUtc As DateTime = File.GetLastWriteTimeUtc(ExcelFilePath)

            Response.Clear()
            Response.Buffer = True
            Response.ContentType = "text/calendar"
            Response.ContentEncoding = Encoding.UTF8

            ' ================================================================
            ' FIX: Change "attachment" to "inline" for Outlook subscription
            ' This makes Outlook treat it as a live calendar subscription
            ' with working reminders and auto-update
            ' ================================================================
            Response.AddHeader("Content-Disposition", "inline; filename=AcademicCalendar.ics")

            ' Cache control for auto-update
            Response.AddHeader("Cache-Control", "no-cache, no-store, must-revalidate, max-age=0")
            Response.AddHeader("Pragma", "no-cache")
            Response.AddHeader("Expires", "-1")
            Response.AddHeader("Last-Modified", excelLastModifiedUtc.ToString("R"))
            Response.AddHeader("ETag", """" & excelLastModifiedUtc.Ticks.ToString() & """")
            Response.AddHeader("X-Content-Type-Options", "nosniff")
            Response.AddHeader("Content-Length", fileBytes.Length.ToString())

            Response.Cache.SetCacheability(HttpCacheability.NoCache)
            Response.Cache.SetNoStore()
            Response.Cache.SetExpires(DateTime.UtcNow.AddMinutes(-1))
            Response.Cache.SetRevalidation(HttpCacheRevalidation.AllCaches)

            Response.BinaryWrite(fileBytes)

            Context.ApplicationInstance.CompleteRequest()

        Catch ex As Exception

            Response.Clear()
            Response.ContentType = "text/plain"
            Response.Write("Error creating calendar feed: " & ex.Message)
            Context.ApplicationInstance.CompleteRequest()

        End Try

    End Sub

    Private Function ReadExcelEvents() As DataTable

        Dim cleanTable As New DataTable()

        cleanTable.Columns.Add("EventTitle", GetType(String))
        cleanTable.Columns.Add("EventDescription", GetType(String))
        cleanTable.Columns.Add("StartDate", GetType(Date))
        cleanTable.Columns.Add("StartTime", GetType(String))
        cleanTable.Columns.Add("EndDate", GetType(Date))
        cleanTable.Columns.Add("EndTime", GetType(String))
        cleanTable.Columns.Add("Location", GetType(String))
        cleanTable.Columns.Add("Category", GetType(String))
        cleanTable.Columns.Add("IsActive", GetType(String))
        cleanTable.Columns.Add("ReminderMinutes", GetType(String))
        cleanTable.Columns.Add("EventID", GetType(String))

        If Not File.Exists(ExcelFilePath) Then
            Throw New FileNotFoundException("Excel file not found. Please put AcademicCalendar.xlsx inside App_Data folder.")
        End If

        Using stream As FileStream = File.Open(ExcelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)

            Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)

                Dim dataSet As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                    .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                        .UseHeaderRow = True
                    }
                })

                If dataSet.Tables.Count = 0 Then
                    Throw New Exception("The Excel file does not contain any sheet.")
                End If

                Dim excelTable As DataTable = Nothing

                If dataSet.Tables.Contains("Sheet1") Then
                    excelTable = dataSet.Tables("Sheet1")
                Else
                    Throw New Exception("Sheet1 was not found in the Excel file.")
                End If

                ValidateRequiredColumns(excelTable)

                Dim excelRowNumber As Integer = 1

                For Each row As DataRow In excelTable.Rows

                    excelRowNumber += 1

                    If IsEmpty(row("EventTitle")) Then
                        Continue For
                    End If

                    Dim eventTitle As String = row("EventTitle").ToString().Trim()

                    If IsEmpty(row("StartDay")) Then
                        Throw New Exception("StartDay is empty for event: " & eventTitle)
                    End If

                    ' Keep the IsActive check to respect "No" status
                    Dim isActiveValue As String = "Yes"

                    If Not IsEmpty(row("IsActive")) Then
                        isActiveValue = row("IsActive").ToString().Trim()
                    End If

                    ' Skip events marked as "No"
                    If isActiveValue.ToLower() <> "yes" Then
                        Continue For
                    End If

                    Dim startDate As Date = ParseExcelDate(row("StartDay"), "StartDay", eventTitle)

                    Dim endDate As Date = startDate

                    If Not IsEmpty(row("EndDay")) Then

                        endDate = ParseExcelDate(row("EndDay"), "EndDay", eventTitle)

                        If endDate < startDate Then
                            Throw New Exception("EndDay cannot be before StartDay for event: " & eventTitle)
                        End If

                    End If

                    Dim startTimeText As String = ""

                    If Not IsEmpty(row("StartTime")) Then
                        startTimeText = ParseExcelTimeText(row("StartTime"), "StartTime", eventTitle)
                    End If

                    Dim endTimeText As String = ""

                    If Not IsEmpty(row("EndTime")) Then
                        endTimeText = ParseExcelTimeText(row("EndTime"), "EndTime", eventTitle)
                    End If

                    If startTimeText <> "" AndAlso endTimeText <> "" Then

                        Dim startDateTime As DateTime = startDate.Date.Add(TimeSpan.Parse(startTimeText))
                        Dim endDateTime As DateTime = endDate.Date.Add(TimeSpan.Parse(endTimeText))

                        If endDateTime <= startDateTime Then
                            Throw New Exception("End date/time must be after Start date/time for event: " & eventTitle)
                        End If

                    End If

                    Dim reminderMinutesText As String = ""

                    If Not IsEmpty(row("ReminderMinutes")) Then

                        reminderMinutesText = row("ReminderMinutes").ToString().Trim()

                        Dim reminderMinutesValue As Integer

                        If Not Integer.TryParse(reminderMinutesText, reminderMinutesValue) Then
                            Throw New Exception("ReminderMinutes must be a number for event: " & eventTitle)
                        End If

                        If reminderMinutesValue < 0 Then
                            Throw New Exception("ReminderMinutes cannot be negative for event: " & eventTitle)
                        End If

                        reminderMinutesText = reminderMinutesValue.ToString()

                    End If

                    Dim newRow As DataRow = cleanTable.NewRow()

                    newRow("EventTitle") = eventTitle
                    newRow("EventDescription") = If(IsEmpty(row("EventDescription")), "", row("EventDescription").ToString().Trim())
                    newRow("StartDate") = startDate
                    newRow("StartTime") = startTimeText
                    newRow("EndDate") = endDate
                    newRow("EndTime") = endTimeText
                    newRow("Location") = If(IsEmpty(row("Location")), "", row("Location").ToString().Trim())
                    newRow("Category") = If(IsEmpty(row("Category")), "", row("Category").ToString().Trim())
                    newRow("IsActive") = "Yes"
                    newRow("ReminderMinutes") = reminderMinutesText
                    newRow("EventID") = "excel-row-" & excelRowNumber.ToString()

                    cleanTable.Rows.Add(newRow)

                Next

            End Using

        End Using

        cleanTable.DefaultView.Sort = "StartDate ASC"
        Return cleanTable.DefaultView.ToTable()

    End Function

    Private Sub ValidateRequiredColumns(excelTable As DataTable)

        Dim requiredColumns As String() = {
            "EventTitle",
            "EventDescription",
            "StartDay",
            "StartTime",
            "EndDay",
            "EndTime",
            "Location",
            "Category",
            "IsActive",
            "ReminderMinutes"
        }

        For Each columnName As String In requiredColumns

            If Not excelTable.Columns.Contains(columnName) Then
                Throw New Exception("Missing Excel column: " & columnName)
            End If

        Next

    End Sub

    Private Function ParseExcelDate(value As Object, fieldName As String, eventTitle As String) As Date

        If value Is Nothing OrElse value Is DBNull.Value OrElse value.ToString().Trim() = "" Then
            Throw New Exception(fieldName & " is empty for event: " & eventTitle)
        End If

        If TypeOf value Is Date Then
            Return Convert.ToDateTime(value)
        End If

        If IsNumeric(value) Then

            Dim numericDate As Double = Convert.ToDouble(value)

            If numericDate > 0 Then
                Return DateTime.FromOADate(numericDate)
            End If

        End If

        Dim textDate As String = value.ToString().Trim()
        Dim parsedDate As Date

        If Date.TryParseExact(textDate, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDate) Then
            Return parsedDate
        End If

        If Date.TryParseExact(textDate, "d-M-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDate) Then
            Return parsedDate
        End If

        If Date.TryParseExact(textDate, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDate) Then
            Return parsedDate
        End If

        If Date.TryParseExact(textDate, "d/M/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDate) Then
            Return parsedDate
        End If

        Throw New Exception("Invalid date in " & fieldName & " for event: " & eventTitle & ". Use dd-mm-yyyy, example: 16-08-2026.")

    End Function

    Private Function ParseExcelTimeText(value As Object, fieldName As String, eventTitle As String) As String

        If value Is Nothing OrElse value Is DBNull.Value OrElse value.ToString().Trim() = "" Then
            Return ""
        End If

        If TypeOf value Is Date Then
            Return Convert.ToDateTime(value).ToString("HH:mm")
        End If

        If IsNumeric(value) Then

            Dim numericTime As Double = Convert.ToDouble(value)

            If numericTime >= 0 AndAlso numericTime < 1 Then
                Return DateTime.FromOADate(numericTime).ToString("HH:mm")
            End If

        End If

        Dim textTime As String = value.ToString().Trim()

        Dim parsedDateTime As DateTime
        Dim parsedTime As TimeSpan

        If DateTime.TryParseExact(textTime, "HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDateTime) Then
            Return parsedDateTime.ToString("HH:mm")
        End If

        If DateTime.TryParseExact(textTime, "H:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDateTime) Then
            Return parsedDateTime.ToString("HH:mm")
        End If

        If TimeSpan.TryParse(textTime, parsedTime) Then

            If parsedTime.Hours >= 0 AndAlso parsedTime.Hours <= 23 AndAlso parsedTime.Minutes >= 0 AndAlso parsedTime.Minutes <= 59 Then
                Return parsedTime.ToString("hh\:mm")
            End If

        End If

        Throw New Exception("Invalid time in " & fieldName & " for event: " & eventTitle & ". Use HH:mm, example: 08:00.")

    End Function

    Private Function BuildICSContent(eventsTable As DataTable) As String

        Dim icsContent As New StringBuilder()
        icsContent.AppendLine("BEGIN:VCALENDAR")
        icsContent.AppendLine("VERSION:2.0")
        icsContent.AppendLine("PRODID:-//Academic Calendar Project//Academic Calendar Feed//EN")
        icsContent.AppendLine("CALSCALE:GREGORIAN")
        icsContent.AppendLine("METHOD:PUBLISH")

        ' Calendar display name in Outlook
        icsContent.AppendLine("X-WR-CALNAME:Academic Calendar")
        icsContent.AppendLine("NAME:Academic Calendar")

        ' Stable calendar identifier for Outlook to track updates
        icsContent.AppendLine("X-WR-RELCALID:academic-calendar-project")

        ' OPTION 2: Floating time with timezone display
        ' Events will display at the same time regardless of user's timezone
        icsContent.AppendLine("X-WR-TIMEZONE:Asia/Beirut")

        ' Auto-update settings for Outlook
        ' Outlook will check for updates every 1 hour
        icsContent.AppendLine("REFRESH-INTERVAL;VALUE=DURATION:PT1H")
        icsContent.AppendLine("X-PUBLISHED-TTL:PT1H")

        ' Additional headers to help with auto-update
        icsContent.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:FREE")
        icsContent.AppendLine("X-MICROSOUTLOOK-CALENDAR:AcademicCalendar")

        ' Add timezone information so Outlook displays correctly
        icsContent.AppendLine("BEGIN:VTIMEZONE")
        icsContent.AppendLine("TZID:Asia/Beirut")
        icsContent.AppendLine("BEGIN:STANDARD")
        icsContent.AppendLine("DTSTART:19700101T000000")
        icsContent.AppendLine("TZOFFSETFROM:+0300")
        icsContent.AppendLine("TZOFFSETTO:+0200")
        icsContent.AppendLine("TZNAME:EET")
        icsContent.AppendLine("END:STANDARD")
        icsContent.AppendLine("BEGIN:DAYLIGHT")
        icsContent.AppendLine("DTSTART:19700330T000000")
        icsContent.AppendLine("TZOFFSETFROM:+0200")
        icsContent.AppendLine("TZOFFSETTO:+0300")
        icsContent.AppendLine("TZNAME:EEST")
        icsContent.AppendLine("END:DAYLIGHT")
        icsContent.AppendLine("END:VTIMEZONE")

        For Each row As DataRow In eventsTable.Rows

            Dim eventTitle As String = row("EventTitle").ToString()
            Dim eventDescription As String = row("EventDescription").ToString()
            Dim eventLocation As String = row("Location").ToString()
            Dim category As String = row("Category").ToString()

            Dim startDate As Date = Convert.ToDateTime(row("StartDate"))
            Dim endDate As Date = Convert.ToDateTime(row("EndDate"))

            Dim startTime As String = row("StartTime").ToString()
            Dim endTime As String = row("EndTime").ToString()

            Dim reminderMinutesText As String = row("ReminderMinutes").ToString().Trim()

            Dim excelLastModifiedUtc As DateTime = File.GetLastWriteTimeUtc(ExcelFilePath)
            Dim sequenceNumber As Integer = CInt(Math.Min(Integer.MaxValue, excelLastModifiedUtc.Subtract(New DateTime(2000, 1, 1)).TotalMinutes))
            Dim uniqueId As String = row("EventID").ToString() & "@academiccalendar"

            icsContent.AppendLine("BEGIN:VEVENT")
            icsContent.AppendLine("UID:" & uniqueId)
            icsContent.AppendLine("DTSTAMP:" & excelLastModifiedUtc.ToString("yyyyMMddTHHmmssZ"))
            icsContent.AppendLine("LAST-MODIFIED:" & excelLastModifiedUtc.ToString("yyyyMMddTHHmmssZ"))
            icsContent.AppendLine("SEQUENCE:" & sequenceNumber.ToString())
            icsContent.AppendLine("STATUS:CONFIRMED")
            icsContent.AppendLine("TRANSP:OPAQUE")

            ' Add the creation date to help with updates
            icsContent.AppendLine("CREATED:" & excelLastModifiedUtc.ToString("yyyyMMddTHHmmssZ"))

            ' OPTION 2: Floating time - No timezone conversion
            ' Events appear at the same time for all users
            If startTime <> "" AndAlso endTime <> "" Then

                Dim startDateTime As DateTime = startDate.Date.Add(TimeSpan.Parse(startTime))
                Dim endDateTime As DateTime = endDate.Date.Add(TimeSpan.Parse(endTime))

                ' FLOATING TIME: No TZID, just the time
                ' This means "10:00 AM" regardless of user's timezone
                icsContent.AppendLine("DTSTART;TZID=Asia/Beirut:" & startDateTime.ToString("yyyyMMddTHHmmss"))
                icsContent.AppendLine("DTEND;TZID=Asia/Beirut:" & endDateTime.ToString("yyyyMMddTHHmmss"))

            Else

                ' All-day events
                Dim icsEndDate As Date = endDate.AddDays(1)
                icsContent.AppendLine("DTSTART;VALUE=DATE:" & startDate.ToString("yyyyMMdd"))
                icsContent.AppendLine("DTEND;VALUE=DATE:" & icsEndDate.ToString("yyyyMMdd"))

            End If

            icsContent.AppendLine("SUMMARY:" & EscapeICS(eventTitle))
            icsContent.AppendLine("LOCATION:" & EscapeICS(eventLocation))
            icsContent.AppendLine("DESCRIPTION:" & EscapeICS(eventDescription))
            icsContent.AppendLine("CATEGORIES:" & EscapeICS(category))

            ' Add reminder if specified
            If reminderMinutesText <> "" Then

                Dim reminderMinutes As Integer = Convert.ToInt32(reminderMinutesText)
                ' Also add Microsoft-specific reminder for better Outlook compatibility
                icsContent.AppendLine("X-MICROSOFT-CDO-REMINDERENABLED:TRUE")
                icsContent.AppendLine("X-MICROSOFT-CDO-REMINDERMINUTESBEFORESTART:" & reminderMinutes.ToString())

                ' Standard VALARM for reminders
                icsContent.AppendLine("BEGIN:VALARM")
                icsContent.AppendLine("TRIGGER;RELATED=START:-PT" & reminderMinutes.ToString() & "M")
                icsContent.AppendLine("ACTION:DISPLAY")
                icsContent.AppendLine("DESCRIPTION:" & EscapeICS("Reminder: " & eventTitle))
                icsContent.AppendLine("END:VALARM")



            End If

            icsContent.AppendLine("END:VEVENT")

        Next

        icsContent.AppendLine("END:VCALENDAR")

        Return icsContent.ToString()

    End Function

    Private Function EscapeICS(value As String) As String

        If value Is Nothing Then
            Return ""
        End If

        Dim cleanValue As String = value.Trim()

        cleanValue = cleanValue.Replace("\", "\\")
        cleanValue = cleanValue.Replace(";", "\;")
        cleanValue = cleanValue.Replace(",", "\,")
        cleanValue = cleanValue.Replace(vbCrLf, "\n")
        cleanValue = cleanValue.Replace(vbCr, "\n")
        cleanValue = cleanValue.Replace(vbLf, "\n")
        cleanValue = cleanValue.Replace("*", "\n")

        Return cleanValue

    End Function

    Private Function CleanUidText(value As String) As String

        If value Is Nothing Then
            Return "event"
        End If

        Dim cleanValue As String = value.ToLower().Trim()

        cleanValue = cleanValue.Replace(" ", "-")
        cleanValue = cleanValue.Replace("/", "-")
        cleanValue = cleanValue.Replace("\", "-")
        cleanValue = cleanValue.Replace(":", "-")
        cleanValue = cleanValue.Replace(";", "-")
        cleanValue = cleanValue.Replace(",", "-")
        cleanValue = cleanValue.Replace(".", "-")
        cleanValue = cleanValue.Replace("*", "-")

        If cleanValue = "" Then
            cleanValue = "event"
        End If

        Return cleanValue

    End Function

    Private Function IsEmpty(value As Object) As Boolean

        Return value Is Nothing OrElse
               value Is DBNull.Value OrElse
               value.ToString().Trim() = ""

    End Function

End Class