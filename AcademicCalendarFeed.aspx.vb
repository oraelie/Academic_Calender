Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Globalization
Imports ExcelDataReader

Public Class AcademicCalendarFeed
    Inherits System.Web.UI.Page

    Private ReadOnly Property ExcelFilePath As String
        Get
            Return Server.MapPath("~/App_Data/AcademicCalendar.xlsm")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

        Try
            Dim dt As DataTable = ReadExcelEvents()
            Dim icsContent As String = BuildICSContent(dt)
            Dim fileBytes As Byte() = Encoding.UTF8.GetBytes(icsContent)

            Response.Clear()
            Response.Buffer = True
            Response.ContentType = "text/calendar"
            Response.ContentEncoding = Encoding.UTF8
            Response.AddHeader("Content-Disposition", "inline; filename=AcademicCalendar.ics")
            Response.AddHeader("Cache-Control", "no-cache, no-store, must-revalidate")
            Response.AddHeader("Pragma", "no-cache")
            Response.AddHeader("Expires", "0")
            Response.BinaryWrite(fileBytes)
            Response.Flush()

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

    Private Function BuildICSContent(eventsTable As DataTable) As String
        Dim icsContent As New StringBuilder()

        icsContent.AppendLine("BEGIN:VCALENDAR")
        icsContent.AppendLine("VERSION:2.0")
        icsContent.AppendLine("PRODID:-//Academic Calendar Project//Academic Calendar Feed//EN")
        icsContent.AppendLine("CALSCALE:GREGORIAN")
        icsContent.AppendLine("METHOD:PUBLISH")
        icsContent.AppendLine("X-WR-CALNAME:Academic Calendar")
        icsContent.AppendLine("X-WR-TIMEZONE:Asia/Beirut")

        For Each row As DataRow In eventsTable.Rows

            Dim eventTitle As String = row("EventTitle").ToString()
            Dim eventDescription As String = row("EventDescription").ToString()
            Dim category As String = row("Category").ToString()

            Dim startDate As Date = Convert.ToDateTime(row("StartDate"))
            Dim endDate As Date = Convert.ToDateTime(row("EndDate"))

            Dim icsEndDate As Date = endDate.AddDays(1)

            Dim uniqueId As String =
                startDate.ToString("yyyyMMdd") & "-" &
                CleanUidText(eventTitle) & "@academiccalendar"

            icsContent.AppendLine("BEGIN:VEVENT")
            icsContent.AppendLine("UID:" & uniqueId)
            icsContent.AppendLine("DTSTAMP:" & DateTime.UtcNow.ToString("yyyyMMddTHHmmssZ"))
            icsContent.AppendLine("DTSTART;VALUE=DATE:" & startDate.ToString("yyyyMMdd"))
            icsContent.AppendLine("DTEND;VALUE=DATE:" & icsEndDate.ToString("yyyyMMdd"))
            icsContent.AppendLine("SUMMARY:" & EscapeICS(eventTitle))
            icsContent.AppendLine("DESCRIPTION:" & EscapeICS(eventDescription))
            icsContent.AppendLine("CATEGORIES:" & EscapeICS(category))
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
        Return value Is Nothing OrElse value Is DBNull.Value OrElse value.ToString().Trim() = ""
    End Function

End Class