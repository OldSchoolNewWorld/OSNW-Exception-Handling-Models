Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Partial Class MainWindow

    Private Sub Do_Window_Initialized(sender As Object, e As EventArgs)

        ' DEV: The outer protective wrapper goes here. It ensures that
        ' unhandled exceptions, either local or in a subsequent call, are
        ' captured.
        Try
            ' DEV: The major intended operation goes here.

            ' DEV: Expected-safe operations go here.
            ' Argument checking.

            '        Throw New System.Exception("Exception forced in " &
            '            "Do_Window_Initialized expected-safe operation.")

            ' DEV: An inner protective wrapper goes here. Exceptions can be
            ' captured and dealt with where there is a better indication of
            ' where things went wrong. Multiple risky operations can be
            ' wrapped separately to further limit the scope of where a problem
            ' occured or to have different reactions in place.
            Try
                ' DEV: A risky operation goes here.

                '            Throw New System.Exception("Exception forced in " &
                '                "Do_Window_Initialized INNER wrapper.")
            Catch CaughtEx As System.Exception
                ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.

                ' Respond to an anticipated exception.
                Dim CaughtBy As System.Reflection.MethodBase =
                    System.Reflection.MethodBase.GetCurrentMethod
                Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)

                ' Optional rethrow of the caught exception.
                'Throw

            Finally
                ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.

                ' DEV: Do any clean-ups or back-outs here.

            End Try

            ' DEV: Expected-safe operations go here.

        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Dim CaughtBy As System.Reflection.MethodBase =
                System.Reflection.MethodBase.GetCurrentMethod()
            Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)
        End Try
    End Sub ' Do_Window_Initialized

    Private Sub Do_Window_Loaded(sender As Object, e As RoutedEventArgs)

        ' DEV: The outer protective wrapper goes here. It ensures that
        ' unhandled exceptions, either local or in a subsequent call, are
        ' captured.
        Try
            ' DEV: The major intended operation goes here.

            ' DEV: Expected-safe operations go here.
            ' Argument checking.

            ' DEV: An inner protective wrapper goes here. Exceptions can be
            ' captured and dealt with where there is a better indication of
            ' where things went wrong. Multiple risky operations can be
            ' wrapped separately to further limit the scope of where a problem
            ' occured or to have different reactions in place.
            Try
                ' DEV: A risky operation goes here.

                '            Throw New NotImplementedException()

            Catch CaughtEx As System.Exception
                ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.

                ' Respond to an anticipated exception.
                Dim CaughtBy As System.Reflection.MethodBase =
                    System.Reflection.MethodBase.GetCurrentMethod
                Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)

                ' Optional rethrow of the caught exception.
                'Throw

            Finally
                ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.

                ' DEV: Do any clean-ups or back-outs here.

            End Try

            ' DEV: Expected-safe operations go here.

        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Dim CaughtBy As System.Reflection.MethodBase =
                System.Reflection.MethodBase.GetCurrentMethod()
            Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)
        End Try
    End Sub ' Do_Window_Loaded

    Private Sub Do_Window_Closing(sender As Object,
                                  e As ComponentModel.CancelEventArgs)

        ' DEV: The outer protective wrapper goes here. It ensures that
        ' unhandled exceptions, either local or in a subsequent call, are
        ' captured.
        Try
            ' DEV: The major intended operation goes here.

            ' DEV: Expected-safe operations go here.
            ' Argument checking.

            ' DEV: An inner protective wrapper goes here. Exceptions can be
            ' captured and dealt with where there is a better indication of
            ' where things went wrong. Multiple risky operations can be
            ' wrapped separately to further limit the scope of where a problem
            ' occured or to have different reactions in place.
            Try
                ' DEV: A risky operation goes here.

                '            Throw New NotImplementedException()
            Catch CaughtEx As System.Exception
                ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.

                ' Respond to an anticipated exception.
                Dim CaughtBy As System.Reflection.MethodBase =
                    System.Reflection.MethodBase.GetCurrentMethod
                Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)

                ' Optional rethrow of the caught exception.
                'Throw

            Finally
                ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.

                ' DEV: Do any clean-ups or back-outs here.

            End Try

            ' DEV: Expected-safe operations go here.

        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Dim CaughtBy As System.Reflection.MethodBase =
                System.Reflection.MethodBase.GetCurrentMethod()
            Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)
        End Try
    End Sub ' Do_Window_Closing

    Private Sub Do_Window_Closed(sender As Object, e As EventArgs)

        ' DEV: The outer protective wrapper goes here. It ensures that
        ' unhandled exceptions, either local or in a subsequent call, are
        ' captured.
        Try
            ' DEV: The major intended operation goes here.

            ' DEV: Expected-safe operations go here.
            ' Argument checking.

            ' DEV: An inner protective wrapper goes here. Exceptions can be
            ' captured and dealt with where there is a better indication of
            ' where things went wrong. Multiple risky operations can be
            ' wrapped separately to further limit the scope of where a problem
            ' occured or to have different reactions in place.
            Try
                ' DEV: A risky operation goes here.
                '            Throw New NotImplementedException()

            Catch CaughtEx As System.Exception
                ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.

                ' Respond to an anticipated exception.
                Dim CaughtBy As System.Reflection.MethodBase =
                    System.Reflection.MethodBase.GetCurrentMethod
                Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)

                ' Optional rethrow of the caught exception.
                'Throw

            Finally
                ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.

                ' DEV: Do any clean-ups or back-outs here.

            End Try

            ' DEV: Expected-safe operations go here.

        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Dim CaughtBy As System.Reflection.MethodBase =
                System.Reflection.MethodBase.GetCurrentMethod()
            Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)
        End Try
    End Sub ' Do_Window_Closed

    Private Sub Do_Test1Button_Click(
      sender As Object, e As RoutedEventArgs)

        ' DEV: The outer protective wrapper goes here. It ensures that
        ' unhandled exceptions, either local or in a subsequent call, are
        ' captured.
        Try
            ' DEV: The major intended operation goes here.

            ' DEV: Expected-safe operations go here.
            ' Argument checking.

            ' DEV: An inner protective wrapper goes here. Exceptions can be
            ' captured and dealt with where there is a better indication of
            ' where things went wrong. Multiple risky operations can be
            ' wrapped separately to further limit the scope of where a problem
            ' occured or to have different reactions in place.
            Try
                ' DEV: A risky operation goes here.

            Catch CaughtEx As System.Exception
                ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.

                ' Respond to an anticipated exception.
                Dim CaughtBy As System.Reflection.MethodBase =
                    System.Reflection.MethodBase.GetCurrentMethod
                Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)

                ' Optional rethrow of the caught exception.
                'Throw

            Finally
                ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.

                ' DEV: Do any clean-ups or back-outs here.

            End Try

            ' DEV: Expected-safe operations go here.

        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Dim CaughtBy As System.Reflection.MethodBase =
                System.Reflection.MethodBase.GetCurrentMethod()
            Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)
        End Try
    End Sub ' Do_Test1Button_Click

    Private Sub Do_Test2Button_Click(
      sender As Object, e As RoutedEventArgs)

        ' DEV: The outer protective wrapper goes here. It ensures that
        ' unhandled exceptions, either local or in a subsequent call, are
        ' captured.
        Try
            ' DEV: The major intended operation goes here.

            ' DEV: Expected-safe operations go here.
            ' Argument checking.

            ' DEV: An inner protective wrapper goes here. Exceptions can be
            ' captured and dealt with where there is a better indication of
            ' where things went wrong. Multiple risky operations can be
            ' wrapped separately to further limit the scope of where a problem
            ' occured or to have different reactions in place.
            Try
                ' DEV: A risky operation goes here.

            Catch CaughtEx As System.Exception
                ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.

                ' Respond to an anticipated exception.
                Dim CaughtBy As System.Reflection.MethodBase =
                    System.Reflection.MethodBase.GetCurrentMethod
                Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)

                ' Optional rethrow of the caught exception.
                'Throw

            Finally
                ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.

                ' DEV: Do any clean-ups or back-outs here.

            End Try

            ' DEV: Expected-safe operations go here.

        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Dim CaughtBy As System.Reflection.MethodBase =
                System.Reflection.MethodBase.GetCurrentMethod()
            Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)
        End Try
    End Sub ' Do_Test2Button_Click

    Private Sub Do_Test3Button_Click(
      sender As Object, e As RoutedEventArgs)

        ' DEV: The outer protective wrapper goes here. It ensures that
        ' unhandled exceptions, either local or in a subsequent call, are
        ' captured.
        Try
            ' DEV: The major intended operation goes here.

            ' DEV: Expected-safe operations go here.
            ' Argument checking.

            ' DEV: An inner protective wrapper goes here. Exceptions can be
            ' captured and dealt with where there is a better indication of
            ' where things went wrong. Multiple risky operations can be
            ' wrapped separately to further limit the scope of where a problem
            ' occured or to have different reactions in place.
            Try
                ' DEV: A risky operation goes here.

            Catch CaughtEx As System.Exception
                ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.

                ' Respond to an anticipated exception.
                Dim CaughtBy As System.Reflection.MethodBase =
                    System.Reflection.MethodBase.GetCurrentMethod
                Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)

                ' Optional rethrow of the caught exception.
                'Throw

            Finally
                ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.

                ' DEV: Do any clean-ups or back-outs here.

            End Try

            ' DEV: Expected-safe operations go here.

        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Dim CaughtBy As System.Reflection.MethodBase =
                System.Reflection.MethodBase.GetCurrentMethod()
            Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)
        End Try
    End Sub ' Do_Test3Button_Click

End Class ' MainWindow
