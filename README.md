#Overview of Project
This analysis was to produce a view of green stocks performance over a year and be able to compare them between 2017 and 2018. 


#Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

##Stock Performance

The stocks in this data set performed much better in 2017 than in 2018. 

###Performance in 2017
[insert image]

###Performance in 2018

[insert image]

##Run time comparisons
The exection times between the original and refactored scripts were nearly a second difference, with the original script running much faster. 

###Original run times

Script sample:
Find the total volume for the current ticker.
    
        If Cells(j, 1).Value = ticker Then
        
        totalVolume = totalVolume + Cells(j, 8).Value
        
        End If

[insert image]

###Refactored run times 

Script sample:
'3a) Increase volume for current ticker
          
            If Cells(j, 1).Value = tickerIndex Then
            tickerVolumes = tickerVolumes + Cells(j, 8).Value
            End If 

[insert image]


#Summary
##What are the advantages or disadvantages of refactoring code?
The advantages of refactoring code are so that we can be confident that as the data increases the code can continue to handle to the output without interruption. Cleaner code increases the chances the code can be reused no matter the data set size. The cons are that refactoring the original code may lead to a more complex code that is more difficult to scale to a larger data set or may increase the risk of error. 


##How do these pros and cons apply to refactoring the original VBA script?
In determining how long the script took to run between the original code and the refactored code. 
