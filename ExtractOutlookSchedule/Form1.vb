Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Imports System.IO
Imports System.ComponentModel

Public Class Form1
    Dim conn As OleDbConnection

    ' there are form level variables so that there is no
    ' need to pass them down to the sub that writes the data
    ' to the database
    Dim currentSubject As String
    Dim currentCategories As String
    Dim currentOrganizer As String
    Dim currentCreated As String
    Dim currentLastModified As String
    Dim currentContents As String
    Dim currentRecipient As String
    Dim currentAccepted As Boolean
    Dim currentStartDate As Date
    Dim currentEndDate As Date
    Dim currentStartTime As String
    Dim currentEndTime As String
    Dim currentTimeTotal As String
    Dim problemCount As Int16

    Private Sub CreateErrorLogFile(errWhere As String, errMessage As String, errStackTrace As String)

        ' write any errors encountered to a log file to help in diagnostics
        Dim logFileName As String = $"{Application.StartupPath()}\ErrorLog_{DateTime.Today.ToString("dd-MMM-yyyy")}.txt"

        Using sw As StreamWriter = New StreamWriter(File.Open(logFileName, FileMode.Append))
            sw.WriteLine($"{DateTime.Now:f}: Error in {errWhere} code")
            sw.WriteLine(errStackTrace)
            sw.WriteLine(errMessage)
            sw.WriteLine("-------------------------------------------------------------------------------------")
        End Using

    End Sub

    Private Sub CreateAccessDatabase()

        ' create the access database to extract the events in to
        Dim sourceDB As String = Path.Combine(Application.StartupPath, "OutlookSchedule.accdb")
        Dim dbPath As String = Path.Combine(Application.StartupPath, $"{Format(Now, "MMddyyy_HHmm")}_OutlookSchedule.accdb")

        Try
            ' using the prepackages DB - create one for this extract
            File.Copy(sourceDB, dbPath)

            conn = New OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath}")
            conn.Open()

        Catch ex As Exception
            MsgBox("Unable to create the database to store the Outlook events.", MsgBoxStyle.Critical, "Create Database")
            CreateErrorLogFile("CreateAccessDatabase", ex.Message, ex.StackTrace)
            conn = Nothing
        End Try

    End Sub

    Private Sub btnExtract_Click(sender As Object, e As EventArgs) Handles btnExtract.Click

        Dim startDate As Date
        Dim endDate As Date

        Dim objOutlook As Outlook.Application
        Dim objNamespace As Outlook.NameSpace
        Dim objAppointment As Outlook.AppointmentItem
        Dim objRecurrence As Outlook.RecurrencePattern
        Dim objOccurrence As Outlook.AppointmentItem
        Dim objAppointments As Outlook.Items
        Dim objRecipient As Outlook.Recipient

        problemCount = 0

        ' date range to pull appointments for
        startDate = dtpFrom.Value.AddDays(-1)
        endDate = dtpTo.Value.AddDays(1)

        ' each run creates its own database
        CreateAccessDatabase()

        If conn Is Nothing Then
            Exit Sub
        End If

        Try
            ' open outlook
            objOutlook = New Outlook.Application
            objNamespace = objOutlook.GetNamespace("MAPI")

            ' added to pull recurring appointments
            objAppointments = objNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar).Items
            objAppointments.Sort("[Start]")
            objAppointments.IncludeRecurrences = True

            ' set the date range we want to pull calendar information for
            objAppointment = objAppointments.Find($"[Start] > '{startDate.ToString("MM/dd/yyyy")}' and [Start] < '{endDate.ToString("MM/dd/yyyy")}'")

            If objAppointment Is Nothing Then
                MsgBox("No scheduled events found for the specificied dates.", MsgBoxStyle.Information, "Extract Schedule")
            End If

            Do While TypeName(objAppointment) <> "Nothing"

                ' these are constant for every iteration of the recipients + days
                currentSubject = objAppointment.Subject

                currentCategories = objAppointment.Categories

                currentOrganizer = objAppointment.Organizer
                currentCreated = objAppointment.CreationTime
                currentLastModified = objAppointment.LastModificationTime
                currentContents = objAppointment.Body

                currentStartDate = Format(objAppointment.Start, "MM/dd/yyyy")
                currentEndDate = Format(objAppointment.End, "MM/dd/yyyy")

                ' simple a visual reference that the extraction is happening
                lbLog.Items.Add($"Extracting '{currentSubject}' by {currentOrganizer}")
                lbLog.Items.Add($"      When: {currentStartDate} - {currentEndDate}")
                Application.DoEvents()

                ' loop through any recipients assigned to the meeting
                For Each objRecipient In objAppointment.Recipients
                    currentRecipient = objRecipient.Name

                    ' there are 6 response values, here were comparing the status
                    ' to the value of accepted, if it equals then true will be saved to the table
                    ' or if the organizer is also the recipient then it is set as accepted
                    currentAccepted = ((objRecipient.MeetingResponseStatus = Outlook.OlMeetingResponse.olMeetingAccepted) Or (objAppointment.Organizer = objRecipient.Name))


                    If currentStartDate = currentEndDate Then
                        ' this is a one day event, write a record per recipient only

                        If objAppointment.AllDayEvent Then
                            currentStartTime = "00:00"
                            currentEndTime = "24:00"
                            currentTimeTotal = 24
                        Else
                            currentStartTime = Format(objAppointment.Start, "HH:mm")
                            currentEndTime = Format(objAppointment.End, "HH:mm")

                            ' duration is saved in minutes in Outlook - divide it by 60 to store hours
                            currentTimeTotal = objAppointment.Duration / 60
                        End If

                        ' call the procedure that write the record to the table
                        WriteRecordToTable()
                    Else
                        ' this spans multiple days, loop until all dates assigned to the meeting are written
                        Do While currentStartDate < Format(objAppointment.End.AddDays(1), "MM/dd/yyyy")

                            If objAppointment.AllDayEvent Then
                                currentStartTime = "00:00"
                                currentEndTime = "24:00"
                                currentTimeTotal = 24
                            Else
                                currentStartTime = Format(objAppointment.Start, "HH:mm")
                                currentEndTime = Format(objAppointment.End, "HH:mm")
                                currentTimeTotal = objAppointment.Duration / 60
                            End If

                            ' writing a record for each day - so the start and end date will be the same
                            currentEndDate = currentStartDate

                            ' call the procedure that write the record to the table
                            WriteRecordToTable()

                            currentStartDate = currentStartDate.AddDays(1)
                        Loop

                        ' reset the starting date so the other recipients are written to the database
                        currentStartDate = Format(objAppointment.Start, "MM/dd/yyyy")
                    End If
                Next ' recipient

                objAppointment = objAppointments.FindNext
            Loop

        Catch ex As Exception
            CreateErrorLogFile("btnExtract", ex.Message, ex.StackTrace)
            MsgBox("Unable to extract Outlook schedule, check the log file for detailed information", MsgBoxStyle.Critical, "Extract Schedule")
        End Try

        ' quit outlook and remove each object from memory
        If Not objOutlook Is Nothing Then objOutlook.Quit()

        objOccurrence = Nothing
        objRecurrence = Nothing
        objAppointment = Nothing
        objNamespace = Nothing
        objOutlook = Nothing

        conn.Close()
        conn = Nothing

        If problemCount > 0 Then
            MsgBox($"There was {problemCount} issue(s) creating record for some events, there are marked with '!!**!!' in the log.", MsgBoxStyle.Information, "Extract Schedule")
        End If

    End Sub

    Private Sub WriteRecordToTable()

        ' create the insert query to write the data to the database
        Dim insertSQL As String
        insertSQL = "Insert into tblCalendarEvents (" _
            & "[Subject], [Catagories], [From], [To], [Created], [Modified], [Accepted], [Date Start], [Date End], [Time Start], " _
            & "[Time End], [Time Total], [Contents], [RecordCreationDateTime]) Values ( "
        insertSQL += $"'{currentSubject}','{currentCategories}','{currentOrganizer}','{currentRecipient}',#{currentCreated}#, "
        insertSQL += $"#{currentLastModified}#, {currentAccepted}, '{currentStartDate}', '{currentEndDate}', '{currentStartTime}', "
        insertSQL += $"'{currentEndTime}', '{currentTimeTotal}', '{currentContents}', #{Now}#)"

        Dim problemMark As String
        Dim cmd As New OleDb.OleDbCommand(insertSQL, conn)
        Dim rdr As OleDbDataReader
        rdr = cmd.ExecuteReader()

        If rdr.RecordsAffected = 0 Then
            ' there was an issue inserting the data into the table
            problemMark = "!!**!!"
            problemCount += 1
        Else
            problemMark = ""
        End If

        ' simple a visual reference that imports are happening
        lbLog.Items.Add($"{problemMark}  Recipient: {currentRecipient}")
        lbLog.Items.Add($"      Start: {currentStartDate} {currentStartTime}")
        lbLog.Items.Add($"      End  : {currentEndDate} {currentEndTime}")
        lbLog.Items.Add($"   Accepted: {currentAccepted}")
        Application.DoEvents()

        rdr.Close()
        cmd.Dispose()

    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles MyBase.Closing

        ' if there is an object in the connection variable then check if it is still open
        If Not conn Is Nothing Then
            If conn.State = ConnectionState.Open Then
                ' if it is then close it
                conn.Close()
            End If
            ' make the var nothing
            conn = Nothing
        End If

    End Sub
End Class
