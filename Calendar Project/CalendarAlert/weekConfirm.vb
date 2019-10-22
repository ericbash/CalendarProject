Public Class WeekConfirm

    Public Sub WeekConfirm_Load(sender As Object, e As EventArgs)
        Me.Text = AlertForm.calendarName

        Dim time As Date = Today
        time = time.AddHours(5)
        time = time.AddMinutes(30)

        'Make table starting at 5:30 AM to 11:00 PM
        For i As Integer = 0 To 33
            Me.DataGridView1.Rows.Add()
            Me.DataGridView1.Rows.Item(i).HeaderCell.Value = time.TimeOfDay.ToString
            time = time.AddMinutes(30)
        Next i

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Will go back to alertForm if more users need to be processed, Or back to begining if none are left
        Me.Hide()
        If AlertForm.userIndex < PickUserForm.passedUsers.Count - 1 Then
            AlertForm.AlertForm_reload(sender, e)
            AlertForm.Show()
        Else
            AlertForm.userIndex = 0
            PickUserForm.Show()
        End If
    End Sub


    'Reloads this form
    Public Sub Reload_WeekConfirm(sender As Object, e As EventArgs)
        Me.Controls.Clear() 'removes all the controls on the form
        InitializeComponent()
        WeekConfirm_Load(sender, e)
    End Sub


    Private Sub WeekConfirm_FormClosing(ByVal sender As Object, ByVal e As EventArgs) Handles Me.FormClosed
        AlertForm.Close()
    End Sub
End Class