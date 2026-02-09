Namespace My

    ' The following events are availble for MyApplication:
    ' 
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication

        Private Sub MyApplication_UnhandledException(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.UnhandledExceptionEventArgs) Handles Me.UnhandledException
            If Microsoft.VisualBasic.MsgBox("An unexpected exception has occurred. The program cannot proceed." & vbCrLf & "Do you want to see the stack?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "SQL Data Manager") = MsgBoxResult.Yes Then Microsoft.VisualBasic.MsgBox(e.Exception.Message & vbCrLf & e.Exception.Source & vbCrLf & e.Exception.StackTrace, MsgBoxStyle.OkOnly, "SQL Data Manager")
        End Sub
    End Class

End Namespace

