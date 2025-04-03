Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Class MainWindow

    ' REF: Best practices for exceptions
    ' https://learn.microsoft.com/en-us/dotnet/standard/exceptions/best-practices-for-exceptions

#Region "Exception Model App Events"
    ' DEV: These events are intended as part of the model.

    ''' <summary>
    ''' Occurs when this <c>Window</c> is initialized. Backing fields can
    ''' usually be set after arriving here. See
    ''' <see cref="System.Windows.FrameworkElement.Initialized"/>.
    ''' </summary>
    Private Sub Window_Initialized(sender As Object, e As EventArgs) _
        Handles Me.Initialized

        Try
            Me.Do_Window_Initialized(sender, e)
        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Dim CaughtBy As System.Reflection.MethodBase =
                System.Reflection.MethodBase.GetCurrentMethod()
            Me.ShowExceptionMessageBox(CaughtBy, CaughtEx, sender, e)
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

        Try
            Me.Do_Window_Loaded(sender, e)
        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Dim CaughtBy As System.Reflection.MethodBase =
                System.Reflection.MethodBase.GetCurrentMethod()
            Me.ShowExceptionMessageBox(CaughtBy, CaughtEx, sender, e)
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

        Try
            Me.Do_Window_Closing(sender, e)
        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Dim CaughtBy As System.Reflection.MethodBase =
                System.Reflection.MethodBase.GetCurrentMethod()
            Me.ShowExceptionMessageBox(CaughtBy, CaughtEx, sender, e)
        End Try
    End Sub ' Window_Closing

    ''' <summary>
    ''' Occurs when the window is about to close. See
    ''' <see cref="System.Windows.Window.Closed"/>.
    ''' </summary>
    Private Sub Window_Closed(sender As Object, e As EventArgs) _
        Handles Me.Closed

        Try
            Do_Window_Closed(sender, e)
        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Dim CaughtBy As System.Reflection.MethodBase =
                System.Reflection.MethodBase.GetCurrentMethod()
            Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)
        End Try
    End Sub ' Window_Closed

    Private Sub CloseButton_Click(sender As Object, e As RoutedEventArgs) _
        Handles CloseButton.Click

        Try
            ' DEV: Expected-safe operations go here. A simple operation that is
            ' not expected to have problems can be left here, without the inner
            ' exception handling process.
            Me.Close()
        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Dim CaughtBy As System.Reflection.MethodBase =
                System.Reflection.MethodBase.GetCurrentMethod()
            Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)
        End Try
    End Sub ' CloseButton_Click

#End Region ' "Exception Model App Events"

#Region "Example Risky App Routines"

    Private Function GetSqrt(input As System.Double) As System.Double
        If input < 0.0 Then
            Throw New System.ArgumentException(
                $"Cannot take the square root of negative value '{input}'.")
        End If
        Return System.Math.Sqrt(input)
    End Function

    Private Sub ProcessDice(ByVal die1 As System.Int32, ByVal die2 As System.Int32)

        ' Argument checking.
        If die1 < 1 OrElse die1 > 6 Then
            Dim NewEx As New System.ArgumentOutOfRangeException(
                $"Invalid value {die1} received for 'die1'")
            NewEx.Data.Add("die1", die1)
            NewEx.Data.Add("die2", die2)
            Throw NewEx
        End If
        If die2 < 1 OrElse die2 > 6 Then
            Dim NewEx As New System.ArgumentOutOfRangeException(
                $"Invalid value {die2} received for 'die2'")
            NewEx.Data.Add("die1", die1)
            NewEx.Data.Add("die2", die2)
            Throw NewEx
        End If

        '
        ' Go ahead and continue the process based on the valid dice.
        '

    End Sub


#End Region ' "Example Risky App Routines"

#Region "Example App Events"
    ' DEV: These events are not intended as part of the model. They are included
    ' to support operation of the example.

    ''' <summary>
    ''' Catches an unhandled exception in a called routine, which bubbles up to
    ''' the outer layer of this calling routine.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OuterButton_Click(
        sender As Object, e As RoutedEventArgs) _
        Handles OuterButton.Click

        Try

            Dim Input As System.Double = -7
            Dim Output As System.Double

            ' DEV: An inner protective wrapper is omitted here.
            ' DEV: A risky operation goes here.
            Output = Me.GetSqrt(Input)
            System.Windows.MessageBox.Show(
                $"The square root of {Input} is {Output}.", "Success!",
                System.Windows.MessageBoxButton.OK)

        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Dim CaughtBy As System.Reflection.MethodBase =
                System.Reflection.MethodBase.GetCurrentMethod()
            Me.ShowExceptionMessageBox(CaughtBy, CaughtEx, sender, e)
        End Try
    End Sub ' OuterButton_Click

    ''' <summary>
    ''' Demonstrates the full structured approach. Catches an unhandled
    ''' exception in a called routine, which bubbles up to the inner layer of
    ''' this calling routine. The inner layer can handle the exception with a
    ''' good understanding of which activity experienced the exception.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub InnerButton_Click(
        sender As Object, e As RoutedEventArgs) _
        Handles InnerButton.Click

        Try

            Dim Input As System.Double = -7
            Dim Output As System.Double

            Try
                Output = Me.GetSqrt(Input)
                System.Windows.MessageBox.Show(
                    $"The square root of {Input} is {Output}.", "Success!",
                    System.Windows.MessageBoxButton.OK)
            Catch CaughtEx As System.Exception
                ' Respond to an anticipated exception.
                Dim CaughtBy As System.Reflection.MethodBase =
                    System.Reflection.MethodBase.GetCurrentMethod
                Me.ShowExceptionMessageBox(CaughtBy, CaughtEx, sender, e)
            End Try

        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Dim CaughtBy As System.Reflection.MethodBase =
                System.Reflection.MethodBase.GetCurrentMethod()
            Me.ShowExceptionMessageBox(CaughtBy, CaughtEx, sender, e)
        End Try
    End Sub ' InnerButton_Click

    Private Sub DiceButton_Click(
        sender As Object, e As RoutedEventArgs) _
        Handles DiceButton.Click

        ' DEV: The outer protective wrapper goes here. It ensures that
        ' unhandled exceptions, either local or in a subsequent call, are
        ' captured.
        Try
            ' DEV: The major intended operation goes here.

            ' DEV: Expected-safe operations go here.
            ' Argument checking.
            ' Setups.

            ' DEV: An inner protective wrapper goes here. Exceptions can be
            ' captured and dealt with where there is a better indication of
            ' where things went wrong. Multiple risky operations can be
            ' wrapped separately to further limit the scope of where a problem
            ' occured or to have different reactions in place.
            Try
                Me.ProcessDice(3, 7)
            Catch CaughtEx As System.Exception

                Dim CaughtBy As System.Reflection.MethodBase =
                    System.Reflection.MethodBase.GetCurrentMethod
                If CaughtEx.Data.Count > 0 Then
                    ' Data is included.
                    Dim DataMessage As System.String =
                        MainWindow.GetDataMessage(CaughtEx)
                    Me.ShowExceptionMessageBox(CaughtBy, CaughtEx, DataMessage)
                Else
                    Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)
                End If

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
    End Sub ' DiceButton_Click

#End Region ' "Example App Events"

End Class ' MainWindow
