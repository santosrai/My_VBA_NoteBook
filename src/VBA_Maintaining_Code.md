## Standard Way of Starting 
 * Comments on Functions or methods
```
'---------------------------------------------------------------------------------------
'Purpose:
'Inputs:
'Return:
'---------------------------------------------------------------------------------------
```

 * Comments on Main Class
 ```
 '------------------------------------------------------------------------------------------------
' Revision History:
'-----------------
' Rev       Date(yyyy/mm/dd)    Version     Description
' **************************************************************************************
' 1          2019/03/12            v1.0      Initial Release
'
' 2          2019/03/26            v2.0     - User Interface Setting for Move File Setting
'                                           - Handle invalid value in userform
'                                           - Code Refactoring
'                                           - File Name is changed  yyyymmddhhmm -> yyyymmddhhmmss
'                                           - Unit test Integration
'-------------------------------------------------------------------------------------------------
```
* Error Handling
```
On Error GoTo ErrLabel
  '// Task

Exit_Handler:   
    '//handling exit way
    Exit Sub     
ErrLabel:   
    MsgBox "エラーが発生しました" & vbCr & vbCr & _
            Err.Description, vbExclamation, "要確認"       
Resume Exit_Handler

```

## Constants should be Capitalized
 eg. FILENAME = ""
 
## Indent your code 


## Refactor the code into simplest form ([https://refactoring.guru](https://refactoring.guru))

- Simplify Conditional Statements

     * Using Template Method (Decompose Conditional)
      - Before Refactor 
        ```
        $data = [];
        if ($code1 == 'A' && $code2 == '01' && in_array($code3, ['1', '3', '5'])) {
            // something
        } elseif ($code1 == 'A' && $code2 == '02' && in_array($code3, ['2', '4', '6'])) {
            // Something
        } elseif (...)
        }
        ```
         - After Refactor 
         
         ```
        $codes = [$code1, $code2, $code3];
        if ($this->isPatternA($codes)) {
             // something
        } elseif ($this->isPatternB($codes)) {
             // Something
        } elseif (...)
        }
    
        ```
   * Replace Nested Conditional with Guard Clauses
      - Before Refactor 
        ```
        If (condition1) Then
            // something
        ElseIf （condition2） Then
            // something
        ElseIf （condition3） Then
            // something
        End If
        ```
         - After Refactor 
         
         ```
        If （condition1） Then
            //something
        Exit Sub
        End If
        
        If （condition2） Then
            //something
        Exit Sub
        End If
        
        If (condition3) Then
            //something
        End If
        ```
    * Replace Constructor [jetbrains.com](https://www.jetbrains.com/help/idea/refactoring-source-code.html)
        - with Factory Method: this refactoring lets you hide a constructor and replace it with a static method which returns a new instance of a class. 
    
        - with Builder :  this refactoring helps hide a constructor, replacing its usages with the references to a newly generated builder class, or to an existing builder class.
        
