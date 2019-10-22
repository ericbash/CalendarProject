Imports Microsoft.Office.Interop

Public Class AlertForm
    Public catchedUserName As String

    Dim counter As Integer

    Dim restrictedcalendarItems As Outlook.Items
    Dim Appointment As Outlook.AppointmentItem = Nothing

    Dim myStart As Date = Today
    Const dateRange As Integer = 4
    Dim dayDiff As Integer = Today.DayOfWeek - DayOfWeek.Sunday
    Dim sunday As Date = Today.AddDays(-dayDiff)
    Dim myEnd As Date = DateAdd("d", dateRange, myStart)

    Private Sub PickUser_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Console.WriteLine("Started calander read")

        Dim otkApp As Outlook.Application
        Dim otkNameSpace As Outlook.NameSpace
        Dim calendar As Outlook.MAPIFolder
        Dim calendarItems As Outlook.Items
        PickUserForm.Hide()
        ProgressBar1.Maximum = 336

        Dim strRestriction As String

        otkApp = New Outlook.Application
        otkNameSpace = otkApp.GetNamespace("MAPI")
        calendar = otkNameSpace.Folders("SharePoint Lists").Folders("IT Section - IT Vacation Calendar")
        calendarItems = calendar.Items


        calendarItems.IncludeRecurrences = True
        calendarItems.Sort("[Start]")
        strRestriction = "[Start] <= '" & myEnd & "' AND [End] >= '" & sunday & "'"
        restrictedcalendarItems = calendarItems.Restrict(strRestriction)

        counter = 0
        CheckedListBox1.CheckOnClick = True
        For Each Appointment In restrictedcalendarItems
            If (Appointment.Subject.Contains(catchedUserName)) Then
                CheckedListBox1.Items.Insert(counter, Appointment.Subject)
                counter += 1
            End If
        Next

        Console.WriteLine("Finshed calander read")

        otkApp = Nothing
        otkNameSpace = Nothing
        calendarItems = Nothing
        Appointment = Nothing

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Console.WriteLine("Started sheet read")

        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWorkbook As Excel.Workbook = xlApp.Workbooks.Open("P:\PR15600\Support\EmployeeDefinitions\" & catchedUserName & "SupportHours.csv")
        Dim xWorksheet As Excel.Worksheet = xlWorkbook.Sheets(1)

        Dim xlRange As Excel.Range = xWorksheet.Range("B5:H52")
        Dim startDay As Date = sunday
        Dim endDay As Date = startDay.AddMinutes(29)


        For i As Integer = 1 To 7
            For j As Integer = 1 To 48
                ProgressBar1.Increment(1)

                If xlRange.Cells(j, i).value2 = "Y" Then
                    counter = 0
                    For Each Appointment In restrictedcalendarItems
                        If Appointment.Subject.Contains(catchedUserName) Then
                            If CheckedListBox1.GetItemCheckState(counter) Then
                                Dim result1 = DateTime.Compare(Appointment.Start, endDay)
                                Dim result2 = DateTime.Compare(Appointment.End, startDay)
                                If result1 <= 0 And result2 > 0 Then
                                    xlRange.Cells(j, i).value2 = "N"
                                End If
                            End If
                            counter += 1
                        End If
                    Next Appointment
                End If
                startDay = startDay.AddMinutes(30)
                endDay = endDay.AddMinutes(30)
            Next j
        Next i

        Console.WriteLine("Finished sheets read")


        xlWorkbook.Save()
        xlWorkbook.Close()

        xlApp = Nothing
        xlWorkbook = Nothing
        xWorksheet = Nothing
        xlRange = Nothing

        PickUserForm.Close()
    End Sub
End Class