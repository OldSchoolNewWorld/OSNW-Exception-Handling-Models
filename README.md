# OSNW-Exception-Handling-Models

This project is a WPF application model to test or demonstrate approaches to 
exception handling. It includes remarks regarding the applicability of certain 
implementation choices.

### Coding Notes

"Option Explicit On", "Option Strict On", and "Option Infer Off" are set in the 
code to make it clear what is being done and to make it easier to research the 
code elements. Fully qualified references are used in much of the code, 
including the use of "Me".

Some of the XML comments are targeted at developers and are left in place so 
that they are visible to IntelliSense while creating a new 
application/assembly. Items marked "DEV:" are intended for a developer using 
the model, not to be included in assemblies based on these models. They can be 
left in place, deleted, suppressed by adding another apostrophe, or suppressed 
from output by the compiler via "Generate XML documentation file".

External research references are marked "REF:". Those are links to research 
done while looking for code samples or detailed explanations of properties and 
methods.

The model includes regions. Those regions may appear to be overkill for the 
simplistic example application but can provide organization for a more complex 
application. In a more complex application, the region content may be worth 
moving to separate modules.

As with regions, the use of subroutines may appear to be overkill for the 
simplistic example application but can provide value for a more complex 
application. In a complex implementation, detailed content may be worth moving 
to separate regions or modules. Subroutines can be used to minimize the code 
shown in a group of event handlers. The call to a known-good subroutine is 
something that can easily be stepped over while debugging. Due to the cost of 
setting up a subroutine call, calls to even small subroutines should probably 
be avoided if that is likely to happen in a large loop. When expected to be 
used in a loop, it may be better to bring the code into the calling routine.