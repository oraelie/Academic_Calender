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
        If Not IsPostBack Then
            CurrentViewMode = "List"
            CurrentCategory = "All"

            SetDefaultMonthFromExcel()
            LoadPage()
        End If
    End Sub

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
        cleanTable.Columns.Add("EndDate", GetType(Date))
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

                Dim excelTable As DataTable = dataSet.Tables(0)

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

                    Dim newRow As DataRow = cleanTable.NewRow()

                    newRow("EventTitle") = eventTitle
                    newRow("EventDescription") = If(IsEmpty(row("EventDescription")), "", row("EventDescription").ToString().Trim())
                    newRow("StartDate") = startDate
                    newRow("EndDate") = endDate
                    newRow("Category") = row("Category").ToString().Trim()
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
            "EndDay",
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
        Return value Is Nothing OrElse value Is DBNull.Value OrElse value.ToString().Trim() = ""
    End Function

    Private Function ParseExcelDate(value As Object, fieldName As String, eventTitle As String) As Date
        If value Is Nothing OrElse value Is DBNull.Value OrElse value.ToString().Trim() = "" Then
            Throw New Exception(fieldName & " is empty for event: " & eventTitle)
        End If

        If TypeOf value Is Date Then
            Return Convert.ToDateTime(value)
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

        Throw New Exception("Invalid date in " & fieldName & " for event: " & eventTitle & ". Use dd-mm-yyyy, example: 02-07-2026.")
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
            Dim monthKey As String = startDate.ToString("yyyy-MM")
            Dim category As String = row("Category").ToString()
            Dim title As String = Server.HtmlEncode(row("EventTitle").ToString())

            If monthKey <> currentMonthKey Then

                If currentMonthKey <> "" Then
                    html.Append("</div>")
                End If

                currentMonthKey = monthKey

                html.Append("<div class='monthly-list-card'>")

                html.Append("<div class='monthly-list-header'>")
                html.Append("<span class='monthly-name'>" & startDate.ToString("MMMM") & "</span>")
                html.Append("<span class='monthly-year'>" & startDate.ToString("yyyy") & "</span>")
                html.Append("</div>")

                html.Append(BuildHolidayStrip(eventsTable, startDate.Month, startDate.Year))

            End If

            If category.ToLower() <> "holidays" Then

                html.Append("<div class='monthly-list-row'>")

                html.Append("<div class='monthly-date'>")
                html.Append(FormatListDate(startDate, endDate))
                html.Append("</div>")

                html.Append("<div class='monthly-dot-cell'>")
                html.Append("<span class='dot " & GetDotClass(category) & "'></span>")
                html.Append("</div>")

                html.Append("<div class='monthly-title'>")
                html.Append(title)
                html.Append("</div>")

                html.Append("<div class='monthly-category'>")
                html.Append("<span class='category-badge " & GetCategoryBadgeClass(category) & "'>")
                html.Append(GetCategoryLabel(category))
                html.Append("</span>")
                html.Append("</div>")

                html.Append("</div>")

            End If

        Next

        If currentMonthKey <> "" Then
            html.Append("</div>")
        End If

        Return html.ToString()
    End Function

    Private Function BuildHolidayStrip(eventsTable As DataTable, monthNumber As Integer, yearNumber As Integer) As String
        Dim html As New StringBuilder()
        Dim hasHoliday As Boolean = False

        For Each row As DataRow In eventsTable.Rows

            Dim startDate As Date = Convert.ToDateTime(row("StartDate"))
            Dim endDate As Date = Convert.ToDateTime(row("EndDate"))
            Dim category As String = row("Category").ToString()

            If startDate.Month = monthNumber AndAlso startDate.Year = yearNumber AndAlso category.ToLower() = "holidays" Then

                If Not hasHoliday Then
                    hasHoliday = True
                    html.Append("<div class='holiday-strip'>")
                    html.Append("<span class='holiday-title'>HOLIDAYS</span>")
                End If

                Dim title As String = Server.HtmlEncode(row("EventTitle").ToString())

                html.Append("<span class='holiday-pill'>")
                html.Append(FormatListDate(startDate, endDate) & " – " & title)
                html.Append("</span>")

            End If

        Next

        If hasHoliday Then
            html.Append("</div>")
        End If

        Return html.ToString()
    End Function

    Private Function FormatListDate(startDate As Date, endDate As Date) As String
        If startDate = endDate Then
            Return startDate.ToString("MMM d")
        End If

        If startDate.Month = endDate.Month AndAlso startDate.Year = endDate.Year Then
            Return startDate.ToString("MMM d") & "–" & endDate.Day.ToString()
        End If

        Return startDate.ToString("MMM d") & "–" & endDate.ToString("MMM d")
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

        Dim filter As String = "StartDate >= #" & firstDay.ToString("MM/dd/yyyy") & "# AND StartDate <= #" & lastDay.ToString("MM/dd/yyyy") & "#"

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
                    html.Append("<div class='day-number'><span class='today-number'>" & currentDate.Day.ToString() & "</span></div>")
                Else
                    html.Append("<div class='day-number " & dayClass & "'>" & currentDate.Day.ToString() & "</div>")
                End If

                For Each row As DataRow In eventsTable.Rows

                    Dim eventDate As Date = Convert.ToDateTime(row("StartDate"))

                    If eventDate.Date = currentDate.Date Then

                        Dim category As String = row("Category").ToString()
                        Dim eventCss As String = GetEventClass(category)
                        Dim title As String = Server.HtmlEncode(row("EventTitle").ToString())

                        html.Append("<div class='calendar-event " & eventCss & "'>")
                        html.Append(title)
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

End Class