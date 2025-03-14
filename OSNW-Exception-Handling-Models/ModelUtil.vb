Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System.Reflection

Partial Class MainWindow
    ' DEV: These routines are intended as part of the model.

#Region "Basic sub/function model"

    ' Private Sub SomeSub()
    '     ' DEV: The outer protective wrapper goes here. It ensures that
    '     ' unhandled exceptions, either local or in a subsequent call, are
    '     ' captured.
    '     Try
    '         ' DEV: The major intended operation goes here.
    ' 
    '         ' DEV: Expected-safe operations go here.
    '         ' Argument checking.
    '         ' Setups.
    ' 
    '         ' DEV: An inner protective wrapper goes here. Exceptions can be
    '         ' captured and dealt with where there is a better indication of
    '         ' where things went wrong. Multiple risky operations can be
    '         ' wrapped separately to further limit the scope of where a problem
    '         ' occured or to have different reactions in place.
    '         Try
    '             ' DEV: A risky operation goes here.
    ' 
    '         Catch CaughtEx As System.Exception
    '             ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.
    ' 
    '             ' Respond to an anticipated exception.
    '             Dim CaughtBy As System.Reflection.MethodBase =
    '                 System.Reflection.MethodBase.GetCurrentMethod
    '             Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)
    ' 
    '             ' Optional rethrow of the caught exception.
    '             'Throw
    ' 
    '         Finally
    '             ' BC30030 - Try must have at least one 'Catch' or a 'Finally'.
    '
    '             ' DEV: Do any clean-ups or back-outs here.
    '
    '         End Try
    ' 
    '         ' DEV: Expected-safe operations go here.
    ' 
    '     Catch CaughtEx As System.Exception
    '         ' Report the unexpected exception.
    '         Dim CaughtBy As System.Reflection.MethodBase =
    '             System.Reflection.MethodBase.GetCurrentMethod()
    '         Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)
    '     End Try
    ' End Sub

#End Region ' "Basic sub/function model"

#Region "Exception message box"

    ''' <summary>
    ''' Extracts the <c>Data</c> entries from the specified <c>Exception</c>.
    ''' </summary>
    ''' <param name="caughtEx">Specifies the <c>Exception</c> from which to
    ''' extract the <c>Data</c> pairs.</param>
    ''' <returns>A sting containing the data.</returns>
    ''' <remarks>
    ''' DEV: This can be modified to report the data in any chosen form.
    ''' </remarks>
    Private Shared Function GetDataMessage(ByVal caughtEx As System.Exception) _
        As System.String

        Dim IsFirstDataPair As System.Boolean = True
        Dim SBuilder As New System.Text.StringBuilder
        For Each OneDictionaryEntry As _
            System.Collections.DictionaryEntry In caughtEx.Data

            If Not IsFirstDataPair Then
                SBuilder.Append("; ")
            End If
            IsFirstDataPair = False

            SBuilder.Append($"{OneDictionaryEntry.Key}")
            SBuilder.Append($"={OneDictionaryEntry.Value}")

        Next
        Return SBuilder.ToString
    End Function ' GetDataMessage

    ''' <summary>
    ''' Reports an invalid call to one of the
    ''' <c>ShowExceptionMessageBox(&lt;varies&gt;)</c> implementations.
    ''' </summary>
    ''' <param name="paramName">Specifies the name of the parameter that was
    ''' invalid.</param>
    ''' <param name="reason">Specifies the reason for the rejection.</param>
    ''' <remarks>This is for invalid calls to
    ''' <c>ShowExceptionMessageBox(&lt;varies&gt;)</c>, not for generic invalid
    ''' procedure calls.</remarks>
    Private Sub ShowExceptionArgNotice(ByVal paramName As System.String,
                                       ByVal reason As System.String)

        Dim CaptionStr As System.String = "Invalid ShowExceptionMessageBox"
        Dim IntroDetails As System.String =
            "An invalid exception notice was requested."

        ' Construct and show the notice.
        Dim ShownDetail As System.String = System.String.Concat(IntroDetails,
            System.Environment.NewLine, System.Environment.NewLine, reason)
        System.Windows.MessageBox.Show(Me, ShownDetail, CaptionStr,
                                       System.Windows.MessageBoxButton.OK,
                                       System.Windows.MessageBoxImage.Error)

    End Sub ' ShowExceptionArgNotice

    ''' <summary>
    ''' Provides a consistent generic appearance for messages.
    ''' </summary>
    ''' <param name="captionStr">Specifies the caption to show on the 
    ''' <c>MessageBox</c>.</param>
    ''' <param name="introDetails">Specifies a summary of the exception.</param>
    ''' <param name="techDetails">Specifies detailed information about the 
    ''' exception.</param>
    ''' <remarks>
    ''' This is mainly intended for exceptions caught in the outer layer of the
    ''' model. It provides high-level detail regarding where to look for
    ''' problems. It can also be used by the inner layer if specific information
    ''' is provided in <paramref name="techDetails"/>.
    ''' </remarks>
    Private Sub ShowExceptionNotice(ByVal captionStr As System.String,
        ByVal introDetails As System.String, ByVal techDetails As System.String)

        ' Construct and show the notice.
        Dim ShownDetail As System.String = System.String.Concat(introDetails,
            System.Environment.NewLine, System.Environment.NewLine, techDetails)
        System.Windows.MessageBox.Show(Me, ShownDetail, captionStr,
                                       System.Windows.MessageBoxButton.OK,
                                       System.Windows.MessageBoxImage.Error)
    End Sub ' ShowExceptionNotice

    ''' <summary>
    ''' Shows details about an <c>Exception</c> that was caught by an event for
    ''' an <c>Object</c>.
    ''' </summary>
    ''' <param name="caughtEx">Specifies the <c>Exception</c> that was
    ''' caught.</param>
    ''' <param name="sender">Specifies the <c>Object</c> that sent the
    ''' event.</param>
    ''' <param name="e">Specifies arguments that are associated with the
    ''' event.</param>
    ''' <remarks>
    ''' A <see cref="System.Windows.RoutedEventArgs"/> contains information that
    ''' a <see cref="System.EventArgs"/> does not have. <see
    ''' cref="System.EventArgs"/> has more than 30 derived types that may
    ''' provide even more specific information. See <see
    ''' cref="ShowExceptionMessageBox(System.Exception, Object,
    ''' System.EventArgs)"/> for information common across exception and event
    ''' types.
    ''' </remarks>
    Private Sub ShowExceptionMessageBox(
        ByVal caughtBy As System.Reflection.MethodBase,
        ByVal caughtEx As System.Exception,
        ByVal sender As Object,
        ByVal e As System.Windows.RoutedEventArgs)

        ' Argument checking.
        If caughtEx Is Nothing Then
            Me.ShowExceptionArgNotice(NameOf(caughtEx),
                $"'{NameOf(caughtEx)}' cannot be 'Nothing'/'Null'.")
            Exit Sub ' Early exit.
        End If
        Dim CaughtByName As System.String = If(caughtBy Is Nothing,
            $"Unspecified '{NameOf(caughtBy)}'", caughtBy.Name)
        Dim SenderName As System.String = If(sender Is Nothing,
            $"Unspecified '{NameOf(sender)}'", sender.ToString)
        Dim EText As System.String = If(e Is Nothing,
            "Unspecified RoutedEventArgs", e.RoutedEvent.ToString)

        ' Unique information in System.Windows.RoutedEventArgs that can be
        ' examined to determine the cause of an exception and where it occurred.

        ' Gets or sets a value that indicates the present state of the event
        ' handling for a routed event as it travels the route.
        Dim EHandled As System.Boolean = e.Handled

        ' Gets the original reporting source as determined by pure hit testing,
        ' before any possible Source adjustment by a parent class.
        Dim EOriginalSource As Object = e.OriginalSource

        ' Gets or sets the RoutedEvent associated with this RoutedEventArgs
        ' instance.
        Dim ERoutedEvent As System.Windows.RoutedEvent = e.RoutedEvent

        ' Gets or sets a reference to the object that raised the event.
        Dim ESource As Object = e.Source

        ' Common System.Exception information that can be examined to determine
        ' the cause of an exception and where it occurred.
        Dim CaughtExTargetSite As System.Reflection.MethodBase =
            caughtEx.TargetSite
        Dim CaughtExBaseException As System.Exception =
            caughtEx.GetBaseException

        ' Gather information of interest.
        Dim IntroDetails As System.String =
            $"An exception was caught in '{CaughtByName}'" &
            $", with sender='{sender}'" & $" and message '{caughtEx.Message}'."
        Dim TechDetails As System.String =
            $"The initial cause is {CaughtExBaseException}."

        ' Construct and show the notice.
        Dim Caption As System.String = "Routed Event Failure"
        Me.ShowExceptionNotice(Caption, IntroDetails, TechDetails)

    End Sub ' ShowExceptionMessageBox

    ''' <summary>
    ''' Shows details about an <c>Exception</c> that was caught by an event for
    ''' an <c>Object</c>.
    ''' </summary>
    ''' <param name="caughtEx">Specifies the <c>Exception</c> that was
    ''' caught.</param>
    ''' <param name="sender">Specifies the <c>Object</c> that sent the
    ''' event.</param>
    ''' <param name="e">Specifies arguments that are associated with the
    ''' event.</param>
    ''' <remarks>
    ''' A <see cref="System.ComponentModel.CancelEventArgs"/> contains
    ''' information that a <see cref="System.EventArgs"/> does not have. See
    ''' <see cref="ShowExceptionMessageBox(System.Exception, Object,
    ''' System.EventArgs)"/> for information common across exception and event
    ''' types.
    ''' </remarks>
    Private Sub ShowExceptionMessageBox(
        ByVal caughtBy As System.Reflection.MethodBase,
        ByVal caughtEx As System.Exception,
        ByVal sender As Object,
        ByVal e As System.ComponentModel.CancelEventArgs)

        ' Argument checking.
        If caughtEx Is Nothing Then
            Me.ShowExceptionArgNotice(NameOf(caughtEx),
                $"'{NameOf(caughtEx)}' cannot be 'Nothing'/'Null'.")
            Exit Sub ' Early exit.
        End If
        Dim CaughtByName As System.String = If(caughtBy Is Nothing,
            $"Unspecified '{NameOf(caughtBy)}'", caughtBy.Name)
        Dim SenderName As System.String = If(sender Is Nothing,
            $"Unspecified '{NameOf(sender)}'", sender.ToString)
        Dim EText As System.String = If(e Is Nothing,
            "Unspecified CancelEventArgs", e.ToString)

        ' Unique information in System.ComponentModel.CancelEventArgs that can
        ' be examined to determine the cause of an exception and where it occurred.

        ' Gets or sets a value indicating whether the event should be canceled.
        Dim ECancel As System.Boolean = e.Cancel

        ' Common Exception information that can be examined to determine the
        ' cause of an exception and where it occurred.
        Dim CaughtExTargetSite As System.Reflection.MethodBase =
            caughtEx.TargetSite
        Dim CaughtExBaseException As System.Exception =
            caughtEx.GetBaseException

        ' Gather information of interest.
        Dim IntroDetails As System.String =
            $"An exception was caught in '{CaughtByName}'" &
            $", with sender='{sender}'" & $" and message '{caughtEx.Message}'."
        Dim TechDetails As System.String =
            $"The initial cause is {CaughtExBaseException}."

        ' Construct and show the notice.
        Dim Caption As System.String = "Cancel Event Failure"

        Me.ShowExceptionNotice(Caption, IntroDetails, TechDetails)

    End Sub ' ShowExceptionMessageBox

    ''' <summary>
    ''' Shows details about an <c>Exception</c> that was caught by the
    ''' <c>Initialized</c> event for a <c>Window</c>.
    ''' </summary>
    ''' <param name="caughtEx">Specifies the <c>Exception</c> that was
    ''' caught.</param>
    ''' <param name="sender">Specifies the <c>Object</c> that sent the
    ''' <c>Initialized</c> event.</param>
    ''' <param name="e">Specifies arguments that are associated with the
    ''' event.</param>
    ''' <remarks>
    ''' <see cref="System.EventArgs"/> is a base class for other event types.
    ''' </remarks>
    Private Sub ShowExceptionMessageBox(
        ByVal caughtBy As System.Reflection.MethodBase,
        ByVal caughtEx As System.Exception,
        ByVal sender As Object, ByVal e As System.EventArgs)

        ' Argument checking.
        If caughtEx Is Nothing Then
            Me.ShowExceptionArgNotice(NameOf(caughtEx),
                $"'{NameOf(caughtEx)}' cannot be 'Nothing'/'Null'.")
            Exit Sub ' Early exit.
        End If
        Dim CaughtByName As System.String = If(caughtBy Is Nothing,
            $"Unspecified '{NameOf(caughtBy)}'", caughtBy.Name)
        Dim SenderName As System.String = If(sender Is Nothing,
            $"Unspecified '{NameOf(sender)}'", sender.ToString)
        Dim EText As System.String = If(e Is Nothing,
            "Unspecified EventArgs", e.ToString)

        ' The following are examples of information in System.Exception and
        ' System.EventArgs that can be examined to determine the cause of an
        ' exception and where it occurred.

        ' Gets a collection of key/value pairs that provide additional
        ' user-defined information about the exception. Data can be updated when
        ' an exception occurs. Parameters that were passed to a routine can be
        ' added to help identify edge cases, etc. that can cause a problem.
        Dim CaughtExData As System.Collections.IDictionary = caughtEx.Data

        ' Structured - Another exception. Returns the Exception that is the root
        ' cause of one or more subsequent exceptions.
        Dim CaughtExBaseException As System.Exception =
            caughtEx.GetBaseException

        ' The type of exception. Gets the runtime type of the current instance.
        Dim CaughtExType As Type = caughtEx.GetType

        ' Gets the name of the type of exception.
        Dim CaughtExTypeName As System.String = CaughtExType.Name

        ' Gets the full name of the type of exception.
        Dim CaughtExTypeFullname As System.String = CaughtExType.FullName

        ' Structured - Another exception. Gets the Exception instance that
        ' caused the current exception.
        Dim CaughtExInnerException As System.Exception = caughtEx.InnerException

        ' Gets a message that describes the current exception.
        Dim CaughtExMessage As System.String = caughtEx.Message

        ' Gets the name of the application or the object that causes the error.
        Dim CaughtExSource As System.String = caughtEx.Source

        ' Gets a string representation of the immediate frames on the call
        ' stack.
        Dim CaughtExStackTrace As System.String = caughtEx.StackTrace

        ' Gets the method that throws the current exception.
        Dim CaughtExTargetSite As System.Reflection.MethodBase =
            caughtEx.TargetSite

        ' Creates and returns a string representation of the current exception.
        Dim CaughtExToString As System.String =
            caughtEx.ToString 'System.Exception: Forced exception

        ' Structured. Gets the Type of the current instance.
        Dim EType As Type = e.GetType

        ' Gets the name of the current member.
        Dim ETypeName As System.String = EType.Name

        ' Gets the fully qualified name of the type, including its namespace but
        ' not its assembly.
        Dim ETypeFullName As System.String = EType.FullName ' System.EventArgs

        ' Returns a string that represents the current object.
        Dim ETypeToString As System.String = e.ToString

        ' Gather information of interest.
        Dim IntroDetails As System.String =
            $"An exception was caught in '{CaughtByName}'" &
            $", with sender='{sender}'" & $" and message '{caughtEx.Message}'."
        Dim TechDetails As System.String =
            $"The initial cause is {CaughtExBaseException}."

        ' Construct and show the notice.
        Dim Caption As System.String = "Initialization Event Failure"
        'Dim ShownDetail As System.String = System.String.Concat(IntroDetails,
        '    System.Environment.NewLine, System.Environment.NewLine, TechDetails)
        'System.Windows.MessageBox.Show(Me,
        '    ShownDetail,
        '    Caption, System.Windows.MessageBoxButton.OK,
        '    System.Windows.MessageBoxImage.Error)

        Me.ShowExceptionNotice(Caption, IntroDetails, TechDetails)

    End Sub ' ShowExceptionMessageBox

    ''' <summary>
    ''' Provides a notice of an unexpected exception.
    ''' </summary>
    ''' <param name="caughtBy">Specifies the process in which an unexpected
    ''' exception occurred.</param>
    ''' <param name="caughtEx">Provides the unexpected exception.</param>
    ''' <remarks>
    ''' </remarks>
    Private Sub ShowExceptionMessageBox(
        ByVal caughtBy As System.Reflection.MethodBase,
        ByVal caughtEx As System.Exception,
        ByVal dataMessage As System.String)

        ' Argument checking.
        If caughtEx Is Nothing Then
            Me.ShowExceptionArgNotice(NameOf(caughtEx),
                $"'{NameOf(caughtEx)}' cannot be 'Nothing'/'Null'.")
            Exit Sub ' Early exit.
        End If
        Dim CaughtByName As System.String = If(caughtBy Is Nothing,
            $"Unspecified '{NameOf(caughtBy)}'", caughtBy.Name)

        ' Information that can be examined to determine the cause of an
        ' exception and where it occurred.
        Dim CaughtExTargetSite As System.Reflection.MethodBase =
            caughtEx.TargetSite
        Dim CaughtExBaseException As System.Exception =
            caughtEx.GetBaseException

        ' Gather information of interest.
        Dim IntroDetails As System.String =
            $"'{CaughtByName}' failed with message" &
            $" '{caughtEx.Message}'" &
            $" and data:{dataMessage}."
        Dim TechDetails As System.String =
            $"The initial cause is {CaughtExBaseException}."

        ' Construct and show the notice.
        Dim Caption As System.String = "Process Failure"
        Me.ShowExceptionNotice(Caption, IntroDetails, TechDetails)

    End Sub ' ShowExceptionMessageBox

    ''' <summary>
    ''' Provides a notice of an unexpected exception.
    ''' </summary>
    ''' <param name="caughtBy">Specifies the process in which an unexpected
    ''' exception occurred.</param>
    ''' <param name="caughtEx">Provides the unexpected exception.</param>
    ''' <remarks>
    ''' </remarks>
    Private Sub ShowExceptionMessageBox(
        ByVal caughtBy As System.Reflection.MethodBase,
        ByVal caughtEx As System.Exception)

        ' Argument checking.
        If caughtEx Is Nothing Then
            Me.ShowExceptionArgNotice(NameOf(caughtEx),
                $"'{NameOf(caughtEx)}' cannot be 'Nothing'/'Null'.")
            Exit Sub ' Early exit.
        End If
        Dim CaughtByName As System.String = If(caughtBy Is Nothing,
            $"Unspecified '{NameOf(caughtBy)}'", caughtBy.Name)

        ' Information that can be examined to determine the cause of an
        ' exception and where it occurred.
        Dim CaughtExTargetSite As System.Reflection.MethodBase =
            caughtEx.TargetSite
        Dim CaughtExBaseException As System.Exception =
            caughtEx.GetBaseException

        ' Gather information of interest.
        Dim IntroDetails As System.String =
            $"'{CaughtByName}' failed with message '{caughtEx.Message}'."
        Dim TechDetails As System.String =
            $"The initial cause is {CaughtExBaseException}."

        ' Construct and show the notice.
        Dim Caption As System.String = "Process Failure"

        Me.ShowExceptionNotice(Caption, IntroDetails, TechDetails)

    End Sub ' ShowExceptionMessageBox

#End Region ' "Exception message box"

End Class ' MainWindow