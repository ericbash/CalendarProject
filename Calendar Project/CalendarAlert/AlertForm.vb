Imports System.IO
Imports Microsoft.Office.Interop

Public Class AlertForm

    'Excel is already open at this point
    'Declare needed varibles
    Public xlWorkbook As Excel.Workbook
    Public xWorksheet As Excel.Worksheet
    Public xlRange As Excel.Range

    Dim Appointment As Outlook.AppointmentItem

    Dim counter As Integer

    Public userIndex As Integer = 0

    Dim restrictedcalendarItemsByName As Outlook.Items

    Public calendarName As String
    Dim loginName As String
    Dim onCallUser As String


    Public Sub AlertForm_Load(sender As Object, e As EventArgs)
        'Get on call user
        Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader("P:\PR15600\Support\OnCall.txt")
        Dim line As String = reader.ReadLine()
        onCallUser = line.Substring(line.IndexOf(",") + 1)

        CheckedListBox1.CheckOnClick = True
        ProgressBar1.Maximum = 336

        'Get passed user names
        Dim user As String = PickUserForm.passedUsers(userIndex)
        calendarName = user.Substring(0, user.LastIndexOf(" "))
        loginName = user.Substring(user.LastIndexOf(" ") + 1)
        Me.Text = calendarName

        'These are the three filters for finding events for users
        Dim filterCal = "@SQL=" & Chr(34) & "http://schemas.microsoft.com/mapi/proptag/0x0037001E" & Chr(34) & " ci_phrasematch " & "'" & calendarName & "'" ' This looks like you could change it but it will break if you do.
        Dim filterLog = Chr(34) & "http://schemas.microsoft.com/mapi/proptag/0x0037001E" & Chr(34) & " ci_phrasematch " & "'" & loginName & "'"
        Dim filter1111 = Chr(34) & "http://schemas.microsoft.com/mapi/proptag/0x0037001E" & Chr(34) & " ci_phrasematch " & "'x1111'"
        'This will get all events with the names of the user but exclude the on call x1111 weekly event
        restrictedcalendarItemsByName = (PickUserForm.restrictedcalendarItemsByDate.Restrict(filterCal & " Or " & filterLog & " And Not " & filter1111))

        'Add all events for users in the list to display
        counter = 0
        For Each Appointment In restrictedcalendarItemsByName
            CheckedListBox1.Items.Insert(counter, Appointment.Subject & vbTab & Appointment.Start & " - " & Appointment.End)
            CheckedListBox1.SetItemChecked(counter, True)
            counter += 1
        Next Appointment

        'Check if multiple users were passed
        If PickUserForm.passedUsers.Count > 1 Then
            userIndex += 1
            If CheckedListBox1.Items.Count = 0 Then
                AlertForm_reload(sender, e)         'Reload page if current user has no events
            End If
        End If

        'If passed a single user and that user has no events, disable continue button
        If CheckedListBox1.Items.Count = 0 And PickUserForm.passedUsers.Count = 1 Then
            Confirm.Enabled = False
        End If

        PickUserForm.Hide()
    End Sub


    Private Sub Confirm_Click(sender As Object, e As EventArgs) Handles Confirm.Click
        If CheckedListBox1.CheckedIndices.Count > 0 Then    'Check if any events are checked
            'Get users excel file
            Dim fileName As String = "P:\PR15600\Support\EmployeeDefinitions\" & loginName & "SupportHours.csv"
            File.Delete(fileName)
            File.Copy("P:\PR15600\Support\EmployeeDefinitions-Reset\" & loginName & "SupportHours.csv", fileName)
            xlWorkbook = PickUserForm.xlApp.Workbooks.Open(fileName)
            xWorksheet = xlWorkbook.Sheets(1)
            xlRange = xWorksheet.Range("B5:H52")

            'Preload calendar view
            WeekConfirm.Reload_WeekConfirm(sender, e)

            'Setup date time comparison variables
            Dim startDay As Date = PickUserForm.sunday
            Dim endDay As Date = startDay.AddMinutes(29)
            Dim result1 As Integer
            Dim result2 As Integer

            'Loop through all cells in excel file
            For i As Integer = 1 To 7
                For j As Integer = 1 To 48
                    ProgressBar1.Increment(1)
                    'Check to see if current user is the on call user, and edit excel file correctly
                    If xlRange.Cells(j, i).value2 = "OC" Then
                        If onCallUser = loginName Then
                            xlRange.Cells(j, i).value2 = "Y"
                        Else
                            xlRange.Cells(j, i).value2 = "N"
                        End If
                    End If

                    '"N"'s will be skipped
                    If xlRange.Cells(j, i).value2 = "Y" Then
                        counter = 0
                        For Each Appointment In restrictedcalendarItemsByName       'loop through all events
                            If CheckedListBox1.GetItemCheckState(counter) Then      'Check if they were checked
                                'Check if they exist within the 30 minutes timeframe of a cell
                                result1 = DateTime.Compare(Appointment.Start, endDay)
                                result2 = DateTime.Compare(Appointment.End, startDay)
                                If result1 <= 0 And result2 > 0 Then
                                    xlRange.Cells(j, i).value2 = "N"            'Set to "N" if it does exist withing cell
                                End If
                            End If
                            counter += 1
                        Next Appointment
                    End If
                    'Move date times by 30 minute
                    startDay = startDay.AddMinutes(30)
                    endDay = endDay.AddMinutes(30)
                    If i >= 2 And i < 7 And j >= 12 And j < 46 Then
                        If xlRange.Cells(j, i).value2 = "N" Then
                            WeekConfirm.DataGridView1.Rows(j - 12).Cells(i - 2).Style.BackColor = Color.Red
                        End If
                    End If
                Next j
            Next i
            xlWorkbook.Save()
            xlWorkbook.Close()

            'Move to next form
            WeekConfirm.Show()
            WeekConfirm.DataGridView1.ClearSelection()
            Me.Hide()
        Else
            AlertForm_reload(sender, e) 'Will move to next user if no checkboxes are checked
        End If
    End Sub


    'Move back to the pickuserform
    Private Sub Back_Click(sender As Object, e As EventArgs) Handles Back.Click
        userIndex = 0
        Me.Hide()
        PickUserForm.Show()
    End Sub


    'This reloads this form
    Public Sub AlertForm_reload(ByVal sender As Object, ByVal e As EventArgs)
        Me.Controls.Clear() 'removes all the controls on the form
        InitializeComponent()
        AlertForm_Load(sender, e)
        Me.Show()
    End Sub


    Private Sub AlertForm_FormClosing(ByVal sender As Object, ByVal e As EventArgs) Handles Me.FormClosed
        PickUserForm.Close()
    End Sub
End Class