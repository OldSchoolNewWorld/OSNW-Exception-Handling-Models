# OSNW Exception Handling Models

This project is a WPF application model to test or demonstrate approaches to 
exception handling. It includes remarks regarding the applicability of certain 
implementation choices.

## Coding Notes

"Option Explicit On", "Option Strict On", and "Option Infer Off" are set in the 
code to make it clear what is being done and to make it easier to research the 
code elements. Fully qualified references used in much of the code, including 
the use of "Me".

Some of the XML comments are targeted at developers and are left in place so 
that they are visible to IntelliSense while creating a new 
application/assembly. Items marked "DEV:" are intended for a developer using 
the model, not for consuming assemblies that use the dialog. They can be left 
in place, deleted, suppressed by adding another apostrophe, or suppressed from 
output by the compiler via "Generate XML documentation file".

External research references are marked "REF:". Those are links to research 
done while looking for code samples or detailed explanations of properties and 
methods.

The model includes regions. Those regions may appear to be overkill for the 
simplistic example application but can provide organization for a more complex 
application. In a more complex application, the region content may be worth 
moving to separate modules.

As with regions, the use of subroutines may appear to be overkill for the 
simplistic example dialog but can provide value for a more complex application. 
In a complex implementation, detailed content may be worth moving to separate 
regions or modules. Subroutines can be used to minimize the code shown in a 
group of event handlers. The call to a known-good subroutine is something that 
can easily be stepped over while debugging. Due to the cost of setting up a 
subroutine call, calls to even small subroutines should probably be avoided if 
that is likely to happen in a large loop. When expected to be used in a loop, 
it may be better to bring the code into the calling routine.

## Exception Handling

The exception handling approach described here uses two levels of detection. An 
outer layer catches unhandled exceptions. It helps to prevent application 
crashes and provides information about the exception that was encountered. An 
inner layer narrows the scope of where an exception is encountered and provides 
a place to do more specific analysis than the outer layer.

### Outer layer

The outer layer provides minimal protection against application crashes due to 
an unexpected, and unhandled, exception. It is just a protective wrapper for 
the real purpose for which the subroutine or function was created.

In some cases, the outer layer may provide enough information to lead the 
developer to the soucre of a problem. In other cases, the inner layer may need 
to be modified to perform better isolation and/or analysis of the problem.

```
    Private Sub SomeSub()
        Try
            ' DEV: The major intended operation goes here.
        Catch CaughtEx As System.Exception
            ' Report the unexpected exception.
            Dim CaughtBy As System.Reflection.MethodBase =
                System.Reflection.MethodBase.GetCurrentMethod()
            Me.ShowExceptionMessageBox(CaughtBy, CaughtEx)
        End Try
    End Sub
```

### Inner layer

The inner layer can provide increased protection against application crashes 
due to an exception and enhanced reporting regarding those exceptions. It 
contains the real purpose for which the subroutine or function was created. One 
or more additional protective wrappers can be applied to code segments that may 
encounter an exception.

Simple code that is not likely to have problems can be left unprotected in the 
inner layer, but still protected by the outer layer. If problems do occur, 
another wrapper can be applied to a code segment that has a problem. For 
example, any argument checking is probably in place to prevent bad arguments 
from causing an exception. That process should probably be omitted from the 
inner protective wrapper.

Some event and exception types provide details specific to those  types. Those 
details can be passed to a custom version of `ShowExceptionMessageBox()` so 
that they can included in the notification.

Only use an inner protective layer if there will be either a `Catch` or 
`Finally` clause. The BC30030 error can be avoided by leaving an unused clause 
in place, but with no active code inside. As long as one is there, the other 
can be omitted.

```
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

        ' Respond to an exception.
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
```

### Full Model

This is the full generic model of the approach used in the sample code.

```
    Private Sub SomeSub()
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

                ' Respond to an exception.
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
    End Sub
```

### Notification

The model includes several versions of `ShowExceptionMessageBox()`. Each 
version has different parameters; they align to the standard event handlers. 
Sample implementations are included for some typical events that are common to 
most WPF applications.

Additional specialized versions could be created that dig into data provided by 
an event that occurred or by the type of `System.Exception` that was thrown. 
That information includes the `System.Exception`, `sender`, various derived 
versions of `System.EventArgs` (`System.Windows.RoutedEventArgs`, 
`System.ComponentModel.CancelEventArgs`, etc.), or arguments passed to a 
routine. Some of the sample code points out values that may be of interest in 
specific cases.

'CaughtEx.Data. can be updated when an exception is caught, to identify 
specific conditions under which an exception is thrown. Examples of that would 
include edge cases and incompatible states. Conditional breakpoints can then be 
used to only stop a debug session at the argument value(s) of interest. See 
[System.Exception.Data](https://learn.microsoft.com/en-us/dotnet/api/system.exception.data?view=net-9.0).
