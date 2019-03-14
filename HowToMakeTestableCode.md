# How To Refactor code to make testable code
## Know which code is difficult to test
  - Code with a lot of conditional behavior which depends on another un-readable code.
  - Different codes responsible for setting the same (global) variable.
  - Code that depends on a long chain of separate evaluations and assignments.
    (i.e voilating the [Single Responsibility Principle (SRP)](https://en.wikipedia.org/wiki/Single_responsibility_principle))
## Always create Single Responsibility Principle Class


## Create Managers and Controllers: Don’t bother trying to pull that responsibility into other individual objects


## Design code with testing in mind 


## Prepare test code for failures 
