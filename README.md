# stock-analysis - VBA Challenge 
Week 2 -VBA 


## Overview of Project
This project is to help Steve analyze a dataset that includes the entire stock market over the last few years. 
Because there are thousands of stocks, so the code I used for analyzing a dozen stocks from 2017, 2018 might not be fast enough. In order to make the code executing time shorter, I refactored code from Module 2 to run the VBA script run faster.


### Purpose
To refactor the VBA code so it can run faster than my pervious code. 


###Results

-  I was able to refactor the code to run the 2017 sheet at 0.1796875 seconds, which is 0.5156250 seconds faster than the previous version, and I was able to make 2018 sheet run time to the same 0.1796875 seconds, which is 0.609375 seconds faster. 
https://raw.githubusercontent.com/yumik20/stock-analysis/main/Resources/VBA_Challenge_2017.png
https://github.com/yumik20/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png
I did notice that even for the same script, sometimes the code runs faster, sometimes it runs slower, but overall the refactored script are faster than before.


### Summary 
1. What are the advantages or disadvantages of refactoring code?
-The advantage of refactoring code is that after refactoring, the script runs faster so it can analyze a large datasets in a shorter time. 
-The Disadvantage were:  it takes a lot of time to test which codes can be deleted or switch. It was hard to figure out whether new variables need to be introduced. 
-It is hard to tell whether the new script will run faster for sure. 

2. How do these pros and cons apply to refactoring the original VBA script?
The pros was that the VBA script does run much faster after refactoring. The cons was that it took a lot of time to test codes. Sometimes when I wrote something not ideal, the VBA constantly showed me overflow and I had to try change the type of data, but still did not work until I restart the coding from scratch. I was not sure why I was asked to introduce a new variable and I did not like it. But after I saw the script runs much faster, I was motivated to even make more changes. 
