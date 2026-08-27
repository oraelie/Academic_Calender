Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Globalization
Imports ExcelDataReader

Public Class AcademicCalendar
    Inherits System.Web.UI.Page

    Private ReadOnly Property ExcelFilePath As String
        Get
            Return Server.MapPath("~/App_Data/AcademicCalendar.xlsm")
        End Get
    End Property

    Private Property CurrentViewMode As String
        Get
            If ViewState("CurrentViewMode") Is Nothing Then
                Return "List"
            End If

            Return ViewState("CurrentViewMode").ToString()
        End Get

        Set(value As String)
            ViewState("CurrentViewMode") = value
        End Set
    End Property

    Private Property CurrentCategory As String
        Get
            If ViewState("CurrentCategory") Is Nothing Then
                Return "All"
            End If

            Return ViewState("CurrentCategory").ToString()
        End Get

        Set(value As String)
            ViewState("CurrentCategory") = value
        End Set
    End Property

    Private Property CurrentMonth As Date
        Get
            If ViewState("CurrentMonth") Is Nothing Then
                Return New Date(Date.Today.Year, Date.Today.Month, 1)
            End If

            Return Convert.ToDateTime(ViewState("CurrentMonth"))
        End Get

        Set(value As Date)
            ViewState("CurrentMonth") = New Date(value.Year, value.Month, 1)
        End Set
    End Property
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

        lnkSubscribeOutlook.NavigateUrl = GetCalendarSubscriptionUrl()

        If Not IsPostBack Then

            CurrentViewMode = "List"
            CurrentCategory = "All"

            SetDefaultMonthFromExcel()
            LoadPage()

        End If

    End Sub
    Private Function GetCalendarSubscriptionUrl() As String

        Dim requestUrl As Uri = Request.Url
        Dim baseUrl As String = requestUrl.GetLeftPart(UriPartial.Authority)

        Dim feedUrl As String = baseUrl & ResolveUrl("~/AcademicCalendarFeed.aspx")

        If feedUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
            Return "webcal://" & feedUrl.Substring("https://".Length)
        End If

        If feedUrl.StartsWith("http://", StringComparison.OrdinalIgnoreCase) Then
            Return "webcal://" & feedUrl.Substring("http://".Length)
        End If

        Return feedUrl

    End Function

    Private Sub SetDefaultMonthFromExcel()

        Try

            Dim dt As DataTable = ReadExcelEvents()

            If dt.Rows.Count > 0 Then

                Dim firstDate As Date = Convert.ToDateTime(dt.Rows(0)("StartDate"))
                CurrentMonth = New Date(firstDate.Year, firstDate.Month, 1)

            Else

                CurrentMonth = New Date(Date.Today.Year, Date.Today.Month, 1)

            End If

        Catch

            CurrentMonth = New Date(Date.Today.Year, Date.Today.Month, 1)

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

        If Not File.Exists(ExcelFilePath) Then
            Throw New FileNotFoundException("Excel file not found. Please put AcademicCalendar.xlsm inside App_Data folder.")
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

                For Each row As DataRow In excelTable.Rows

                    If IsEmpty(row("EventTitle")) Then
                        Continue For
                    End If

                    Dim eventTitle As String = row("EventTitle").ToString().Trim()

                    If IsEmpty(row("StartDay")) Then
                        Throw New Exception("StartDay is empty for event: " & eventTitle)
                    End If

                    Dim isActiveValue As String = "Yes"

                    If Not IsEmpty(row("IsActive")) Then
                        isActiveValue = row("IsActive").ToString().Trim()
                    End If

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
            "IsActive"
        }

        For Each columnName As String In requiredColumns

            If Not excelTable.Columns.Contains(columnName) Then
                Throw New Exception("Missing Excel column: " & columnName)
            End If

        Next

    End Sub

    Private Function IsEmpty(value As Object) As Boolean

        Return value Is Nothing OrElse
               value Is DBNull.Value OrElse
               value.ToString().Trim() = ""

    End Function
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

    Private Sub LoadPage()

        Try

            lblError.Text = ""
            lblError.Visible = False

            pnlListView.Visible = CurrentViewMode = "List"
            pnlCalendarView.Visible = CurrentViewMode = "Calendar"

            btnListView.CssClass = If(CurrentViewMode = "List", "view-btn view-btn-active", "view-btn")
            btnCalendarView.CssClass = If(CurrentViewMode = "Calendar", "view-btn view-btn-active", "view-btn")

            If CurrentViewMode = "List" Then
                LoadListView()
            Else
                LoadCalendarView()
            End If

        Catch ex As Exception

            lblError.Text = ex.Message
            lblError.Visible = True

        End Try

    End Sub

    Private Sub LoadListView()

        Dim dt As DataTable = GetFilteredEventsForList()

        If dt.Rows.Count = 0 Then

            litListEvents.Text = ""
            lblListMessage.Text = "No events found."

        Else

            litListEvents.Text = BuildMonthlyListHtml(dt)
            lblListMessage.Text = ""

        End If

    End Sub

    Private Function GetFilteredEventsForList() As DataTable

        Dim dt As DataTable = ReadExcelEvents()
        Dim view As New DataView(dt)

        If CurrentCategory <> "All" Then
            view.RowFilter = "Category = '" & CurrentCategory.Replace("'", "''") & "'"
        End If

        view.Sort = "StartDate ASC"

        Return view.ToTable()

    End Function

    Private Function BuildMonthlyListHtml(eventsTable As DataTable) As String

        Dim html As New StringBuilder()
        Dim currentMonthKey As String = ""

        For Each row As DataRow In eventsTable.Rows

            Dim startDate As Date = Convert.ToDateTime(row("StartDate"))
            Dim endDate As Date = Convert.ToDateTime(row("EndDate"))

            Dim startTime As String = row("StartTime").ToString()
            Dim endTime As String = row("EndTime").ToString()
            Dim location As String = row("Location").ToString()

            Dim monthKey As String = startDate.ToString("yyyy-MM")
            Dim category As String = row("Category").ToString()
            Dim title As String = FormatTextWithLineBreaks(row("EventTitle").ToString())

            If monthKey <> currentMonthKey Then

                If currentMonthKey <> "" Then
                    html.Append("</div>")
                    html.Append("</div>")
                End If

                currentMonthKey = monthKey

                Dim bodyId As String = "monthBody_" & startDate.ToString("yyyyMM")
                Dim iconId As String = "monthIcon_" & startDate.ToString("yyyyMM")

                Dim isCurrentMonth As Boolean =
                    startDate.Month = Date.Today.Month AndAlso startDate.Year = Date.Today.Year

                Dim defaultDisplay As String = If(isCurrentMonth, "block", "none")
                Dim defaultIcon As String = If(isCurrentMonth, "−", "+")
                Dim defaultState As String = If(isCurrentMonth, "expanded", "collapsed")
                Dim isCurrentMonthText As String = If(isCurrentMonth, "yes", "no")

                html.Append("<div class='monthly-list-card' onclick=""toggleMonthCard('" & bodyId & "', '" & iconId & "')"">")

                html.Append("<div class='monthly-list-header'>")
                html.Append("<div class='monthly-header-title'>")
                html.Append("<span class='monthly-name'>" & startDate.ToString("MMMM") & "</span>")
                html.Append("<span class='monthly-year'>" & startDate.ToString("yyyy") & "</span>")
                html.Append("</div>")

                html.Append("<button type='button' class='collapse-btn' >")
                html.Append("<span id='" & iconId & "'>" & defaultIcon & "</span>")
                html.Append("</button>")

                html.Append("</div>")

                html.Append("<div id='" & bodyId & "' class='monthly-list-body' data-default-state='" & defaultState & "' data-current-month='" & isCurrentMonthText & "' style='display:" & defaultDisplay & ";'>")

            End If

            html.Append("<div class='monthly-list-row'>")

            html.Append("<div class='monthly-date'>")
            html.Append(FormatListDate(startDate, endDate))
            html.Append("</div>")

            html.Append("<div class='monthly-dot-cell'>")
            html.Append("<span class='dot " & GetDotClass(category) & "'></span>")
            html.Append("</div>")

            html.Append("<div class='monthly-title'>")
            html.Append(title)

            If location <> "" Then
                html.Append("<br /><span class='monthly-location'>📍 " & Server.HtmlEncode(location) & "</span>")
            End If

            html.Append("</div>")

            html.Append("<div class='monthly-category'>")
            html.Append("<span class='category-badge " & GetCategoryBadgeClass(category) & "'>")
            html.Append(GetCategoryLabel(category))
            html.Append("</span>")
            html.Append("</div>")

            html.Append("</div>")

        Next

        If currentMonthKey <> "" Then
            html.Append("</div>")
            html.Append("</div>")
        End If

        Return html.ToString()

    End Function
    Private Function FormatListDate(startDate As Date, endDate As Date) As String

        If startDate = endDate Then
            Return startDate.ToString("ddd - MMM d")
        End If

        If startDate.Month = endDate.Month AndAlso startDate.Year = endDate.Year Then
            Return startDate.ToString("ddd - MMM d") & "–" & endDate.ToString("ddd - MMM d")
        End If

        Return startDate.ToString("ddd - MMM d") & "–" & endDate.ToString("ddd - MMM d")

    End Function

    Private Function GetCategoryBadgeClass(category As String) As String

        Select Case category.ToLower()

            Case "exams"
                Return "badge-exams"

            Case "deadlines"
                Return "badge-deadlines"

            Case "registration"
                Return "badge-registration"

            Case "holidays"
                Return "badge-holidays"

            Case "academic"
                Return "badge-academic"

            Case Else
                Return "badge-academic"

        End Select

    End Function

    Private Function GetCategoryLabel(category As String) As String

        Select Case category.ToLower()

            Case "exams"
                Return "EXAM"

            Case "deadlines"
                Return "DEADLINE"

            Case "registration"
                Return "REGISTRATION"

            Case "holidays"
                Return "HOLIDAY"

            Case "academic"
                Return "ACADEMIC"

            Case Else
                Return category.ToUpper()

        End Select

    End Function

    Private Sub LoadCalendarView()

        litCalendarMonth.Text = CurrentMonth.ToString("MMMM")
        litCalendarYear.Text = CurrentMonth.ToString("yyyy")

        Dim dt As DataTable = GetFilteredEventsForCalendar()

        litCalendar.Text = BuildCalendarHtml(dt)

    End Sub

    Private Function GetFilteredEventsForCalendar() As DataTable

        Dim dt As DataTable = ReadExcelEvents()
        Dim view As New DataView(dt)

        Dim firstDay As New Date(CurrentMonth.Year, CurrentMonth.Month, 1)
        Dim lastDay As Date = firstDay.AddMonths(1).AddDays(-1)

        Dim filter As String =
            "StartDate >= #" & firstDay.ToString("MM/dd/yyyy") & "# AND StartDate <= #" & lastDay.ToString("MM/dd/yyyy") & "#"

        If CurrentCategory <> "All" Then
            filter &= " AND Category = '" & CurrentCategory.Replace("'", "''") & "'"
        End If

        view.RowFilter = filter
        view.Sort = "StartDate ASC"

        Return view.ToTable()

    End Function

    Private Function BuildCalendarHtml(eventsTable As DataTable) As String

        Dim html As New StringBuilder()

        Dim firstDayOfMonth As New Date(CurrentMonth.Year, CurrentMonth.Month, 1)

        'Monday start calendar.
        Dim daysBack As Integer = (CInt(firstDayOfMonth.DayOfWeek) + 6) Mod 7
        Dim startCalendarDate As Date = firstDayOfMonth.AddDays(-daysBack)

        html.Append("<table class='calendar-table'>")

        html.Append("<thead>")
        html.Append("<tr>")
        html.Append("<th>MON</th>")
        html.Append("<th>TUE</th>")
        html.Append("<th>WED</th>")
        html.Append("<th>THU</th>")
        html.Append("<th>FRI</th>")
        html.Append("<th>SAT</th>")
        html.Append("<th>SUN</th>")
        html.Append("</tr>")
        html.Append("</thead>")

        html.Append("<tbody>")

        Dim currentDate As Date = startCalendarDate

        For week As Integer = 1 To 6

            html.Append("<tr>")

            For day As Integer = 1 To 7

                html.Append("<td>")

                Dim dayClass As String = ""

                If currentDate.Month <> CurrentMonth.Month Then
                    dayClass = "other-month"
                End If

                If currentDate.Date = Date.Today.Date Then

                    html.Append("<div class='day-number'>")
                    html.Append("<span class='today-number'>" & currentDate.Day.ToString() & "</span>")
                    html.Append("</div>")

                Else

                    html.Append("<div class='day-number " & dayClass & "'>")
                    html.Append(currentDate.Day.ToString())
                    html.Append("</div>")

                End If

                For Each row As DataRow In eventsTable.Rows

                    Dim eventDate As Date = Convert.ToDateTime(row("StartDate"))

                    If eventDate.Date = currentDate.Date Then

                        Dim category As String = row("Category").ToString()
                        Dim eventCss As String = GetEventClass(category)
                        Dim title As String = FormatTextWithLineBreaks(row("EventTitle").ToString())
                        Dim startTime As String = row("StartTime").ToString()
                        Dim endTime As String = row("EndTime").ToString()
                        Dim location As String = row("Location").ToString()

                        html.Append("<div class='calendar-event " & eventCss & "'>")

                        If startTime <> "" AndAlso endTime <> "" Then
                            html.Append("<strong>" & startTime & " - " & endTime & "</strong><br />")
                        ElseIf startTime <> "" Then
                            html.Append("<strong>" & startTime & "</strong><br />")
                        ElseIf endTime <> "" Then
                            html.Append("<strong>Until " & endTime & "</strong><br />")
                        End If

                        html.Append(title)

                        If location <> "" Then
                            html.Append("<br /><span class='calendar-location'>📍 " & Server.HtmlEncode(location) & "</span>")
                        End If

                        html.Append("</div>")

                    End If

                Next

                html.Append("</td>")

                currentDate = currentDate.AddDays(1)

            Next

            html.Append("</tr>")

        Next

        html.Append("</tbody>")
        html.Append("</table>")

        Return html.ToString()

    End Function

    Private Function GetDotClass(category As String) As String

        Select Case category.ToLower()

            Case "exams"
                Return "dot-exams"

            Case "deadlines"
                Return "dot-deadlines"

            Case "registration"
                Return "dot-registration"

            Case "holidays"
                Return "dot-holidays"

            Case "academic"
                Return "dot-academic"

            Case Else
                Return "dot-academic"

        End Select

    End Function

    Private Function GetEventClass(category As String) As String

        Select Case category.ToLower()

            Case "exams"
                Return "event-exams"

            Case "deadlines"
                Return "event-deadlines"

            Case "registration"
                Return "event-registration"

            Case "holidays"
                Return "event-holidays"

            Case "academic"
                Return "event-academic"

            Case Else
                Return "event-academic"

        End Select

    End Function

    Protected Sub btnListView_Click(sender As Object, e As EventArgs) Handles btnListView.Click

        CurrentViewMode = "List"
        LoadPage()

    End Sub

    Protected Sub btnCalendarView_Click(sender As Object, e As EventArgs) Handles btnCalendarView.Click

        CurrentViewMode = "Calendar"
        LoadPage()

    End Sub

    Protected Sub btnPrevMonth_Click(sender As Object, e As EventArgs) Handles btnPrevMonth.Click

        CurrentMonth = CurrentMonth.AddMonths(-1)
        LoadPage()

    End Sub

    Protected Sub btnNextMonth_Click(sender As Object, e As EventArgs) Handles btnNextMonth.Click

        CurrentMonth = CurrentMonth.AddMonths(1)
        LoadPage()

    End Sub

    Protected Sub lnkAll_Click(sender As Object, e As EventArgs) Handles lnkAll.Click

        CurrentCategory = "All"
        LoadPage()

    End Sub

    Protected Sub lnkExams_Click(sender As Object, e As EventArgs) Handles lnkExams.Click

        CurrentCategory = "Exams"
        LoadPage()

    End Sub

    Protected Sub lnkDeadlines_Click(sender As Object, e As EventArgs) Handles lnkDeadlines.Click

        CurrentCategory = "Deadlines"
        LoadPage()

    End Sub

    Protected Sub lnkRegistration_Click(sender As Object, e As EventArgs) Handles lnkRegistration.Click

        CurrentCategory = "Registration"
        LoadPage()

    End Sub

    Protected Sub lnkHolidays_Click(sender As Object, e As EventArgs) Handles lnkHolidays.Click

        CurrentCategory = "Holidays"
        LoadPage()

    End Sub

    Protected Sub lnkAcademic_Click(sender As Object, e As EventArgs) Handles lnkAcademic.Click

        CurrentCategory = "Academic"
        LoadPage()

    End Sub

    Protected Sub btnDownloadICS_Click(sender As Object, e As EventArgs) Handles btnDownloadICS.Click

        Try

            Dim dt As DataTable = ReadExcelEvents()
            Dim icsContent As String = BuildICSContent(dt)
            Dim fileBytes As Byte() = Encoding.UTF8.GetBytes(icsContent)

            Response.Clear()
            Response.Buffer = True
            Response.ContentType = "text/calendar"
            Response.ContentEncoding = Encoding.UTF8
            Response.AddHeader("Content-Disposition", "attachment; filename=AcademicCalendar.ics")
            Response.AddHeader("Content-Length", fileBytes.Length.ToString())
            Response.BinaryWrite(fileBytes)
            Response.Flush()

            Context.ApplicationInstance.CompleteRequest()

        Catch ex As Exception

            lblError.Text = "Error creating Outlook calendar file: " & ex.Message
            lblError.Visible = True

        End Try

    End Sub

    Private Function BuildICSContent(eventsTable As DataTable) As String

        Dim icsContent As New StringBuilder()

        icsContent.AppendLine("BEGIN:VCALENDAR")
        icsContent.AppendLine("VERSION:2.0")
        icsContent.AppendLine("PRODID:-//Academic Calendar Project//Academic Calendar//EN")
        icsContent.AppendLine("CALSCALE:GREGORIAN")
        icsContent.AppendLine("METHOD:PUBLISH")
        icsContent.AppendLine("X-WR-CALNAME:Academic Calendar")
        icsContent.AppendLine("X-WR-TIMEZONE:Asia/Beirut")

        For Each row As DataRow In eventsTable.Rows

            Dim eventTitle As String = row("EventTitle").ToString()
            Dim eventDescription As String = row("EventDescription").ToString()
            Dim eventLocation As String = row("Location").ToString()
            Dim category As String = row("Category").ToString()

            Dim startDate As Date = Convert.ToDateTime(row("StartDate"))
            Dim endDate As Date = Convert.ToDateTime(row("EndDate"))

            Dim startTime As String = row("StartTime").ToString()
            Dim endTime As String = row("EndTime").ToString()

            Dim uniqueId As String =
                startDate.ToString("yyyyMMdd") & "-" &
                CleanUidText(eventTitle) & "@academiccalendar"

            icsContent.AppendLine("BEGIN:VEVENT")
            icsContent.AppendLine("UID:" & uniqueId)
            icsContent.AppendLine("DTSTAMP:" & DateTime.UtcNow.ToString("yyyyMMddTHHmmssZ"))

            If startTime <> "" AndAlso endTime <> "" Then

                Dim startDateTime As DateTime = startDate.Date.Add(TimeSpan.Parse(startTime))
                Dim endDateTime As DateTime = endDate.Date.Add(TimeSpan.Parse(endTime))

                icsContent.AppendLine("DTSTART:" & startDateTime.ToUniversalTime().ToString("yyyyMMddTHHmmssZ"))
                icsContent.AppendLine("DTEND:" & endDateTime.ToUniversalTime().ToString("yyyyMMddTHHmmssZ"))

            Else

                Dim icsEndDate As Date = endDate.AddDays(1)

                icsContent.AppendLine("DTSTART;VALUE=DATE:" & startDate.ToString("yyyyMMdd"))
                icsContent.AppendLine("DTEND;VALUE=DATE:" & icsEndDate.ToString("yyyyMMdd"))

            End If

            icsContent.AppendLine("SUMMARY:" & EscapeICS(eventTitle))
            icsContent.AppendLine("LOCATION:" & EscapeICS(eventLocation))
            icsContent.AppendLine("DESCRIPTION:" & EscapeICS(eventDescription))
            icsContent.AppendLine("CATEGORIES:" & EscapeICS(category))
            icsContent.AppendLine("END:VEVENT")

        Next

        icsContent.AppendLine("END:VCALENDAR")

        Return icsContent.ToString()

    End Function

    Private Function FormatTextWithLineBreaks(value As String) As String

        If value Is Nothing Then
            Return ""
        End If

        Dim safeText As String = Server.HtmlEncode(value.Trim())

        safeText = safeText.Replace("*", "<br />")

        Return safeText

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

End Class