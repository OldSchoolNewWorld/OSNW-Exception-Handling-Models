Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Partial Class MainWindow

    Private Sub Do_Window_Initialized(sender As Object, e As EventArgs)

        ' DEV: Expected-safe operations go here.
        '        Throw New System.Exception("Exception forced in " &
        '            "Do_Window_Initialized expected-safe operation.")

        ' DEV: The inner protective wrapper goes here. Exceptions can be
        ' captured and dealt with where there is a better indication of
        ' what went wrong. Multiple risky operations can be wrapped
        ' separately to further limit the scope of where a problem occured
        ' or to have different reactions in place.
        Try
            ' DEV: A risky operation goes here.
            '            Throw New System.Exception("Exception forced in " &
            '                "Do_Window_Initialized INNER wrapper.")

            ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.
            'Catch CaughtEx As System.Exception

            '    ' Respond to an exception.

            '    ' Optional rethrow of the caught exception
            '    'Throw

            '    ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.
        Finally
            ' DEV: Do any clean-ups or back-outs here.
        End Try

        ' DEV: Expected-safe operations go here.

    End Sub ' Do_Window_Initialized

    Private Sub Do_Window_Loaded(sender As Object, e As RoutedEventArgs)

        ' DEV: Expected-safe operations go here.

        ' DEV: The inner protective wrapper goes here. Exceptions can be
        ' captured and dealt with where there is a better indication of
        ' what went wrong. Multiple risky operations can be wrapped
        ' separately to further limit the scope of where a problem occured
        ' or to have different reactions in place.
        Try
            ' DEV: A risky operation goes here.
            '            Throw New NotImplementedException()

            '    ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.
            'Catch CaughtEx As System.Exception

            '    ' Respond to an exception.

            '    ' Optional rethrow of the caught exception
            '    'Throw

            '    ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.
        Finally
            ' DEV: Do any clean-ups or back-outs here.
        End Try

        ' DEV: Expected-safe operations go here.

    End Sub ' Do_Window_Loaded

    Private Sub Do_Window_Closing(sender As Object,
                                  e As ComponentModel.CancelEventArgs)

        ' DEV: Expected-safe operations go here.

        ' DEV: The inner protective wrapper goes here. Exceptions can be
        ' captured and dealt with where there is a better indication of
        ' what went wrong. Multiple risky operations can be wrapped
        ' separately to further limit the scope of where a problem occured
        ' or to have different reactions in place.
        Try
            ' DEV: A risky operation goes here.
            '            Throw New NotImplementedException()

            '    ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.
            'Catch CaughtEx As System.Exception

            '    ' Respond to an exception.

            '    ' Optional rethrow of the caught exception
            '    'Throw

            '    ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.
        Finally
            ' DEV: Do any clean-ups or back-outs here.
        End Try

        ' DEV: Expected-safe operations go here.

    End Sub ' Do_Window_Closing

    Private Sub Do_Window_Closed(sender As Object, e As EventArgs)

        ' DEV: Expected-safe operations go here.

        ' DEV: The inner protective wrapper goes here. Exceptions can be
        ' captured and dealt with where there is a better indication of
        ' what went wrong. Multiple risky operations can be wrapped
        ' separately to further limit the scope of where a problem occured
        ' or to have different reactions in place.
        Try
            ' DEV: A risky operation goes here.
            '            Throw New NotImplementedException()

            '    ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.
            'Catch CaughtEx As System.Exception

            '    ' Respond to an exception.

            '    ' Optional rethrow of the caught exception
            '    'Throw

            '    ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.
        Finally
            ' DEV: Do any clean-ups or back-outs here.
        End Try

        ' DEV: Expected-safe operations go here.

    End Sub ' Do_Window_Closed

End Class ' MainWindow
