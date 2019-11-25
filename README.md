# vba-test-suite
### Lightweight and easy testing for macros
 
Most macros become legacy code the moment they are written. Tests can fix this. 

By having unit tests for your macro, you can <br>
1. make sure your functions do what you think they should do, <br>
2. make sure your code still works after each change.

The VBA Test Suite enables you to run individual tests and groups of tests in one click. 

---

### Example output (printed to debuging window):

```
#Main_Utils
#test_Quicksort2d
       ....SUCCESS    | 2 equals 2
       ....SUCCESS    | Alpha equals Alpha
       ....SUCCESS    | 1 equals 1
       ....SUCCESS    | Zulu equals Zulu
#test_GetRecordDict
       ....SUCCESS    | Akron equals Akron
       ....SUCCESS    | 165 equals 165
#test_CloneDict
       ....SUCCESS    | 1 equals 1
       ....SUCCESS    | is true
       ....SUCCESS    |  equals 
       ....SUCCESS    | 1 equals 1
       ....SUCCESS    | is true
       ....SUCCESS    |  equals 
#test_GetHeaders
       ....SUCCESS    | 3 equals 3
#test_RowToDict
       ....SUCCESS    | TestVal1 equals TestVal1
       ....SUCCESS    | TestVal2 equals TestVal2
       ....SUCCESS    |  equals 
#test_GetFirstColRg
       ....SUCCESS    | 2 equals 2
--------------------------------
    Main_Utils
--------------------------------
TEST OBJS:  7
TESTS RUN:  17
SUCCESSES:  17
FAILURES :  0
--------------------------------
```
