# stock-analysis

## Overview of Project 
The purpose of this project was to refactor the code that was created in the module 2.   We want to improve the design of the code without changing the output once the code is run.  The code in module 2 and the refactored code should yield the same output.  <br/>


## Results
![VBA_Challenge_2017](https://user-images.githubusercontent.com/33010018/147714693-0a4d48ea-7638-4691-b8ed-6f3c8db49b1b.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/33010018/147714739-83c7bb5f-a34f-4886-a508-5eaf419038c5.png)

For 2018, my refactored code output did not match the output in the original code.  However, my output for 2017 was the same for both the refactored and original code.  I do not understand how this is possible and after reviewing my code I could not figure out the issue.  I didn't want to mess with my code again because I wanted to avoid an overflow error that I kept receiving earlier.  

Based on the results I received, the return for ENPH decreased from 2017 to 2018. Also, my refactored code yielded a faster result than the original. 



## Summary
Advantages of Refactoring Code:<br/>
Refactoring code could be useful to make code look cleaner and neater.  Someone else may need to read or edit your code and this may help give a better understanding of what the code is doing.  Another advantage could be a better performance.  For example, my results were faster with the refactored code.

Disadvantages of Refactoring Code:<br/>
Refactored code could take longer to create than the original.  It can be time consuming to figure out how to make the original code more efficient. Another disadvantage is not receiving the same output as the original.  

Advantages of the Original/Refactored Script:<br/>
I didn't receive any errors while working on the original script.  It was straighforward and did not have as many arrays.  

Disadvantages of the Original/Refactored Script: <br/>
The refactored script kept giving confusing overflow errors for the following lines: tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value and  Cells(4 + 3, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1.  After, reading about overflow errors, I still could not understand what the issue was.  
The original code didn't output the results faster than the refactored.

