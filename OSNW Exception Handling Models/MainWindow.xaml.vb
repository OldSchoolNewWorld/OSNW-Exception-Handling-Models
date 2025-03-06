Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Class MainWindow

    ' REF: Best practices for exceptions
    ' https://learn.microsoft.com/en-us/dotnet/standard/exceptions/best-practices-for-exceptions

#Region "Model Events"
    ' DEV: These events are intended as part of the model.

    ''' <summary>
    ''' Occurs when this <c>Window</c> is initialized. Backing fields can
    ''' usually be set after arriving here. See
    ''' <see cref="System.Windows.FrameworkElement.Initialized"/>.
    ''' </summary>
    Private Sub Window_Initialized(sender As Object, e As EventArgs) _
        Handles Me.Initialized

        ' DEV: The outer protective wrapper goes here. It ensures that
        ' exceptions, either local or in a subsequent call, are captured.
        Try
            ' DEV: The major intended operation goes here.
            Me.Do_Window_Initialized(sender, e)
        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Me.ShowExceptionMessageBox(CaughtEx, sender, e)
        End Try
    End Sub ' Window_Initialized

    ''' <summary>
    ''' Occurs when the <c>Window</c> is laid out, rendered, and ready for
    ''' interaction. Sometimes updates have to wait until arriving here. See
    ''' <see cref="System.Windows.FrameworkElement.Loaded"/>.
    ''' </summary>
    Private Sub Window_Loaded(
        sender As Object, e As RoutedEventArgs) _
        Handles Me.Loaded

        ' DEV: The outer protective wrapper goes here. It ensures that
        ' exceptions, either local or in a subsequent call, are captured.
        Try
            ' DEV: The major intended operation goes here.
            Me.Do_Window_Loaded(sender, e)
        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Me.ShowExceptionMessageBox(CaughtEx, sender, e)
        End Try
    End Sub ' Window_Loaded

    ''' <summary>
    ''' Occurs directly after System.Windows.Window.Close is called, and can be
    ''' handled to cancel window closure. See
    ''' <see cref="System.Windows.Window.Closing"/>.
    ''' </summary>
    Private Sub Window_Closing(
        sender As Object, e As ComponentModel.CancelEventArgs) _
        Handles Me.Closing

        ' DEV: The outer protective wrapper goes here. It ensures that
        ' exceptions, either local or in a subsequent call, are captured.
        Try
            ' DEV: The major intended operation goes here.
            Me.Do_Window_Closing(sender, e)
        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Me.ShowExceptionMessageBox(CaughtEx, sender, e)
        End Try
    End Sub ' Window_Closing

    ''' <summary>
    ''' Occurs when the window is about to close. See
    ''' <see cref="System.Windows.Window.Closed"/>.
    ''' </summary>
    Private Sub Window_Closed(sender As Object, e As EventArgs) _
        Handles Me.Closed

        ' DEV: The outer protective wrapper goes here. It ensures that
        ' exceptions, either local or in a subsequent call, are captured.
        Try
            ' DEV: The major intended operation goes here.
            Do_Window_Closed(sender, e)
        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Me.ShowExceptionMessageBox(CaughtEx)
        End Try
    End Sub ' Window_Closed

    Private Sub CloseButton_Click(sender As Object, e As RoutedEventArgs) _
        Handles CloseButton.Click

        ' DEV: The outer protective wrapper goes here. It ensures that
        ' exceptions, either local or in a subsequent call, are captured.
        Try
            ' DEV: Expected-safe operations go here. A simple operation that is
            ' not expected do have problems can be left here, without the inner
            ' exception handling process.
            Me.Close()
        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Me.ShowExceptionMessageBox(CaughtEx)
        End Try
    End Sub ' CloseButton_Click

#End Region ' "Model Events"

End Class ' MainWindow
