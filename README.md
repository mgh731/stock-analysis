# Overview of Project
This analysis was to produce a view of green stocks performance over a year and be able to compare them between 2017 and 2018. 


# Results

## Stock Performance

The stocks in this data set performed much better in 2017 than in 2018. 

### Performance in 2017
![VBA_Challenge_Output2017](/Resources/VBA_Challenge_Output2017.png)

### Performance in 2018

![VBA_Challenge_Output2018](/Resources/VBA_Challenge_Output2018.png)

## Run time comparisons
The exection times between the original and refactored scripts were nearly a second difference, with the original script running much faster. 

### Original run times

Script sample:
Find the total volume for the current ticker.
    
        If Cells(j, 1).Value = ticker Then
        
        totalVolume = totalVolume + Cells(j, 8).Value
        
        End If

![VBA_Original_2017](/Resources/VBA_Original_2017.png)

![VBA_Original_2018](/Resources/VBA_Original_2018.png)

### Refactored run times 

Script sample:
'3a) Increase volume for current ticker
          
            If Cells(j, 1).Value = tickerIndex Then
            tickerVolumes = tickerVolumes + Cells(j, 8).Value
        
            End If 

![VBA_Challenge_2017](/Resources/VBA_Challenge_2017.png)
![VBA _Challenge_2018](/Resources/VBA _Challenge_2018.png)

# Summary

## What are the advantages or disadvantages of refactoring code?
The advantages of refactoring code are so that we can be confident that as the data increases the code can continue to handle to the output without interruption. Cleaner code increases the chances the code can be reused no matter the data set size. The cons are that refactoring the original code may lead to a more complex code that is more difficult to scale to a larger data set or may increase the risk of error. 


## How do these pros and cons apply to refactoring the original VBA script?
In determining how long the script took to run between the original code and the refactored code. The original code took less time to run than the refactored code.
