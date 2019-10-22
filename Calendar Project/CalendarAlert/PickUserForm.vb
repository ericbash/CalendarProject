Imports System.IO
Imports Microsoft.Office.Interop
Public Class PickUserForm

    Dim myStart As Date = Today
    Const dateRange As Integer = 6                              'Number of days after Sunday to Saturday
    Dim dayDiff As Integer = Today.DayOfWeek - DayOfWeek.Sunday 'Differnce of today and sunday
    Public sunday As Date = Today.AddDays(-dayDiff)             'Set day to sunday date
    Dim myEnd As Date = DateAdd("d", dateRange, myStart)        'Set end date for saturday

    Dim otkApp As Outlook.Application = New Outlook.Application 'Open outlook in background
    Dim otkNameSpace As Outlook.NameSpace = otkApp.GetNamespace("MAPI") 'This is needed. Not sure why
    Dim calendar As Outlook.MAPIFolder = otkNameSpace.Folders("SharePoint Lists").Folders("IT Section - IT Vacation Calendar") 'Get calendar info in ost file
    Dim calendarItems As Outlook.Items = calendar.Items                 'Get all events from calendar
    Public restrictedcalendarItemsByDate As Outlook.Items               'This will be used as a container for events that take place this week

    Public xlApp As Excel.Application = New Excel.Application()         'Open excel in background

    Public passedUsers As New Generic.List(Of String)                   'Used to pass the it(s) name to the next form for use
    Dim users As New Generic.List(Of String)                            'Used as a holder for all users if the "All" checkbox is ticked


    Private Sub PickUser_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Choose what users you would like to edit"

        Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader("P:\PR15600\Support\Calendar\it_users.txt")  'Read txt file of it users 
        reader.ReadLine()                           'skip headers

        While Not reader.EndOfStream                'Read until end of file
            Dim currentRow As String = reader.ReadLine
            currentRow.Trim(" ")                    'Remove spaces before and after line
            If Not currentRow.Equals("") Then       'Make sure no blank lines
                ListBox1.Items.Add(currentRow)      'Add users to listbox
                users.Add(currentRow)               'Add all users to list
            End If
        End While

        reader.Close()

        'Remove all events that do not take place or overlap with this current week
        restrictedcalendarItemsByDate = calendarItems.Restrict("[Start] <= '" & myEnd & "' AND [End] >= '" & sunday & "'")
        restrictedcalendarItemsByDate.IncludeRecurrences = True
        restrictedcalendarItemsByDate.Sort("[Start]")
    End Sub


    Private Sub Continue_Click(sender As Object, e As EventArgs) Handles Button1.Click, ListBox1.DoubleClick
        passedUsers.Clear()                                     'Make sure passedUsers is empty
        If Not IsNothing(ListBox1.SelectedItem) Then            'If a user is selected
            passedUsers.Add(ListBox1.SelectedItem.ToString)     'Pass the selected user and load next form
            AlertForm.AlertForm_reload(sender, e)
        ElseIf CheckBox1.Checked Then                           'If the "All" checkbox is checked
            passedUsers.AddRange(users)                         'Pass all users abnd load next form
            AlertForm.AlertForm_reload(sender, e)
        End If
        'Do nothing if nothing has been selected or checked
    End Sub

    'This will disable the slection of users when the "All" checkbox is checked.
    Private Sub AllButton_Click(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If ListBox1.Enabled Then
            ListBox1.Enabled = False
            ListBox1.ClearSelected()
        Else
            ListBox1.Enabled = True
        End If
    End Sub


    Private Sub AlertForm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        xlApp.Quit()    'Close excel
        'Release all used objects for excel
        ReleaseObject(AlertForm.xlRange)
        ReleaseObject(AlertForm.xWorksheet)
        ReleaseObject(AlertForm.xlWorkbook)
        ReleaseObject(xlApp)
        'Force garbage collection and close program
        GC.Collect()
        GC.WaitForPendingFinalizers()
        SplashScreen1.Close()
    End Sub


    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class