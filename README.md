# My_VBA_NoteBook

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
        - LBound(Arr) would cause a run time error(source:http://www.cpearson.com/excel/isarrayallocated.aspx)\
            so IsError Function help to find array whether array is allocated or not
                
