Public NotInheritable Class SplashScreen1
    Private Sub SplashScreen1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Application title
        If My.Application.Info.Title <> "" Then
            ApplicationTitle.Text = My.Application.Info.Title
        Else
            'If the application title is missing, use the application name, without the extension
            ApplicationTitle.Text = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If
        'Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Revision, My.Application.Info.Version.MinorRevision)

        Version.Text = Version.Text & "2.0"
        Author.Text = Author.Text & "Eric Bash"
    End Sub


    'Once splash is visable to user, start load of PickUserFoam. When that is done, hide the splash screen
    Private Sub Splash_loaded(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        PickUserForm.Show()
        Me.Hide()
    End Sub
End Class
