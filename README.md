# My VBA NoteBook Tips

# 2018
- Default type of variables in VBA is Variant
- Implicit and explicit option variables:
     * Implicit\
      - eg. Dim var1, var2 as Integer\
      - explaination: var2 is declared as Integer but var1 isnot. var1 is implicitly variant as default not Integer
     * Explicit\
      - eg. Dim var1 as Integer &nbsp; _ &nbsp;  Dim var1 as Integer, var2 as Integer\
        &nbsp; Dim var2 as Integer\
      - explaination: We can use as type Modifier with a single Dim
- Parameter ByRef
    - When you use an array as a parameter it cannot use ByVal, it must use ByRef. You can pass the array using ByVal making the parameter a variant.

# 2019
- Types of Arrays in VBA(An Array is said to be allocated if it consumes memory, has valid lower and upper bounds and contain data even if that data is the default values for the data type of the aray, such as empty strings)
    * Static Array
        - set in Dim statement
        - Easy to find Empty or Not
    * Dyanamic Array
        - dimensioned with ReDim statement
        - difficult to find empty with isEmpty Function
        - LBound(Arr) would cause a run time error(source:http://www.cpearson.com/excel/isarrayallocated.aspx)
            so IsError Function help to find array whether array is allocated or not
                
- Why global variables are not preffered in any programming language?\
       - Global variable has global scope. So any function in the program can modify it, and you don’t have any clue who       modified it.\
       - Makes debugging difficult as no clue who modified the value.\
       - Not suitable for multi-threaded program as threads would end up with race condition. To avoid this you may need sync mechanism which would slow down your program.\
       - Reduces readability of source code, as it is difficult to locate where the global variable is defined and where it gets modified/accessed. If you have large number of source files and good number of global variables then good luck in tracking the global variables. You really need to good IDE to make your life easy, still you will end up with bugs.\
       - If the variable occupies large space, then that much of space gets reserved till program exits, irrespective of the program needs it or not.
-  Avoid “.Select”: Try to avoid using the “.Select” method or “.Selection” object, for performance reasons, mostly.
- Encapsulating API Calls
    - Declarations are placed at the top of the module and outside any Subs or Functions
    - Wrappers perform input validation to ensure all parameters are passed as expected
       
- Variable Declaring
    - Dim UserName As String = Dim UserName$
        
- Difference between COM Add-in(CAI) and XLA type Add-in
    - Performance: A CAI is an ActiveX DLL that is compiled into native machine code.
    - Code Security: Using DLL file as code source, CAI is very secured.
    - Multiple Application Support: CAI can be written to support multiple applications
    - Callable Functions support: CAI allow you to call functions in the DLL directly from worksheet cells

- Difference between Type vs Enum
    - Enum : It contains an enumeration of constants, in simple manner, you can group the multiple message in one constant

```
// Without enum
Const msg_Welcome = 1
Const msg_Error = 2
Const msg_Warning = 3
Debug.Print "Message number: " & msg_Welcome


//With enum
Enum msg
    Welcome = 1
    Error = 2
    Warning = 3
End Enum
```
              
    
       
