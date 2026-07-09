Imports System.Data
Imports System.IO
Imports System.Text
Imports ExcelDataReader

Public Class AcademicCalendar
    Inherits System.Web.UI.Page

    Private ReadOnly Property ExcelFilePath As String
        Get
            Return Server.MapPath("~/App_Data/AcademicCalendar.xlsx")
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

                Dim excelTable As DataTable = dataSet.Tables(0)

                ValidateRequiredColumns(excelTable)

                For Each row As DataRow In excelTable.Rows

                    If IsEmpty(row("EventTitle")) Then
                        Continue For
                    End If

                    If IsEmpty(row("StartDate")) Then
                        Continue For
                    End If

                    Dim isActiveValue As String = "Yes"

                    If Not IsEmpty(row("IsActive")) Then
                        isActiveValue = row("IsActive").ToString().Trim()
                    End If

                    If isActiveValue.ToLower() <> "yes" Then
                        Continue For
                    End If

                    Dim startDate As Date = Convert.ToDateTime(row("StartDate"))
                    Dim endDate As Date = startDate

                    If Not IsEmpty(row("EndDate")) Then
                        endDate = Convert.ToDateTime(row("EndDate"))
                    End If

                    Dim newRow As DataRow = cleanTable.NewRow()

                    newRow("EventTitle") = row("EventTitle").ToString().Trim()
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
            "StartDate",
            "EndDate",
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

    Private Function GetFilteredEvents() As DataTable
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
        Dim dt As DataTable = GetFilteredEvents()

        litListMonthTitle.Text = CurrentMonth.ToString("MMMM yyyy")

        Dim displayTable As New DataTable()
        displayTable.Columns.Add("DayText", GetType(String))
        displayTable.Columns.Add("DotClass", GetType(String))
        displayTable.Columns.Add("EventTitle", GetType(String))

        For Each row As DataRow In dt.Rows
            Dim startDate As Date = Convert.ToDateTime(row("StartDate"))
            Dim category As String = row("Category").ToString()

            Dim displayRow As DataRow = displayTable.NewRow()

            displayRow("DayText") = startDate.ToString("MMM d")
            displayRow("DotClass") = "dot " & GetDotClass(category)
            displayRow("EventTitle") = row("EventTitle").ToString()

            displayTable.Rows.Add(displayRow)
        Next

        rptListEvents.DataSource = displayTable
        rptListEvents.DataBind()

        If displayTable.Rows.Count = 0 Then
            lblListMessage.Text = "No events found for this month."
        Else
            lblListMessage.Text = ""
        End If
    End Sub

    Private Sub LoadCalendarView()
        litCalendarMonth.Text = CurrentMonth.ToString("MMMM")
        litCalendarYear.Text = CurrentMonth.ToString("yyyy")

        Dim dt As DataTable = GetFilteredEvents()

        litCalendar.Text = BuildCalendarHtml(dt)
    End Sub

    Private Function BuildCalendarHtml(eventsTable As DataTable) As String
        Dim html As New StringBuilder()

        Dim firstDayOfMonth As New Date(CurrentMonth.Year, CurrentMonth.Month, 1)
        Dim startCalendarDate As Date = firstDayOfMonth.AddDays(-CInt(firstDayOfMonth.DayOfWeek))

        html.Append("<table class='calendar-table'>")
        html.Append("<thead>")
        html.Append("<tr>")
        html.Append("<th>SUN</th>")
        html.Append("<th>MON</th>")
        html.Append("<th>TUE</th>")
        html.Append("<th>WED</th>")
        html.Append("<th>THU</th>")
        html.Append("<th>FRI</th>")
        html.Append("<th>SAT</th>")
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