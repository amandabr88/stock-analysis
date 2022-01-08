# stock-analysis

## Overview of the project

### Background
    
    Steve wants wants to expand the dataset to include the entire stock market over the last few years to do more research for his parents. 
    Refactoring the code is a key part of process. In order to help Steve, we are creating a new version of the macro by refactoring the code and giving to him the possibility to works for thousands of stocks. 

### Purpose

    The purpose of this project is to development a solution code to loop through all the data one time in order to collect the same information that we did in the All Stock Macro. For that, we should refactore the code to make the VBA script more efficient.  


## Results

    1. The stock performance for 2017 is much better in comparison to 2018. As we have flagged by the Red color. 
    
    ![Stock Comparison](/Resources/Stock_performance_comparison.png)

    2. The results for Green Stock file was accurate matching for the original macro in the AllStockAnalysis module activity and the refactored code (VBA_Challenge.xlsm).
    3. The execution time in the refactored code is more efficient in comparison to the original macro (Module activity).

    Point 2 and 3 are exemplified by the images below.

    ![VBA_GreenStock_2017_Module2](/Resources/VBA_GreenStock_2017_Module2.png)
    ![VBA_Challenge_2017_Refactored](/Resources/VBA_Challenge_2017_Refactored.png)
    ![VBA_GreenStock_2018_Module2](/Resources/VBA_GreenStock_2018_Module2.png)
    ![VBA_Challenge_2018_Refactored](/Resources/VBA_Challenge_2018_Refactored.png)
    

    4. Code was optimized by elimination of a nested for loopd. This reduced computation time.

    ![Code Comparison](/Resources/CodeComparison.png)

## Summary

    What are the advantages or disadvantages of refactoring code?
    
    The advantages of refactoring code may include improved the execution time , make the code easier to understand and easier to debug.

    The disadvantages of refactoring code may include more time and money to develop. In case you don't have time to refactored the code, it is not recommended to try.

    How do these pros and cons apply to refactoring the original VBA script?
    For this particular example, the refactored code optimized the execution time. Also, make the code easier to understand by avoiding nested loops.
