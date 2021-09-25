# stock-analysis

## Overview of Project
The purpose of this project was to write efficient VBA code that would allow Steve to be able to easily and quickly analyze stock performance. First we learned how to write code to accomplish the task of calculating and outputting select stock performance metrics from an Excel dataset into a separate Excel sheet. Then we refactored that code to loop through all the data just one time in order to make the VBA script run faster. The refactored script allows Steve to expand the dataset to analyze the entire stock market more efficiently.

## Results

### Stock Performance
The dataset included 12 stocks and how they performed in 2017 and 2018. They had a much better return in 2017 than they did in 2018, as seen below.

![image](https://github.com/JFoArlas/stock-analysis/blob/main/Resources/VBA_Challenge_2017_stock%20list.PNG)
![image](https://github.com/JFoArlas/stock-analysis/blob/main/Resources/VBA_Challenge_2018_stock%20list.PNG)

To make it clear which stocks had a positive vs. a negative return, I used the following For Loop to highlight positive return values in green and negative return values in red.

```
For i = dataRowStart To dataRowEnd

  If Cells(i, 3) > 0 Then 
    Cells(i, 3).Interior.Color = vbGreen
  
  Else
    Cells(i, 3).Interior.Color = vbRed
            
  End If
        
Next i
```

### Execution Times
The execution time for the refactored code was far faster than the original code since it looped through the dataset just once. 

*2017 Original vs. Refactored Execution Times:*

![2017 Original Execution Time](https://github.com/JFoArlas/stock-analysis/blob/main/Resources/VBA_Challenge_2017_original.PNG)
![2017 Refactored Execution Time](https://github.com/JFoArlas/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

*2018 Original vs. Refactored Execution Times:*

![2018 Original Execution Time](https://github.com/JFoArlas/stock-analysis/blob/main/Resources/VBA_Challenge_2018_original.PNG)
![2018 Refactored Execution Time](https://github.com/JFoArlas/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

To capture how long the script would take to run, I started the subroutine by initializing variables for `startTime` and `endTime` and then entered `startTime = timer` after the line of code that prompts the user to input which year the analysis should be run on. 
```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer
```

Then at the end of the subroutine I entered `endTime = Timer` before a `MsgBox` function that would show how long the script took to run.

```
    endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
End Sub
```

## Summary: In a summary statement, address the following questions.
### Advantages & Disadvantages of Refactoring Code
Some advantages of refactoring code are the ability to use fewer steps, less memory, or improve the logic of the code to make it easier for future users to read. According to users on [stack overflow](https://stackoverflow.com/questions/43983284/what-are-the-advantages-and-disadvantages-of-refactoring-code-smell-in-software), some disadvantages are that refactoring code can be time consuming, require a lot of retesting, and can be risky on large applications or when the existing code does not have proper test cases.

### Pros & Cons of Refactoring this original VBA Script
The pros outweigh the cons in this scenario, since the time it took to refactor the original code was worth the highly improved execution time. 
