# Stock-Analysis

## Overview of Project

Steve's parents are excited about investing in alternate energy stock DAQO New Energy Corp (ticker 'DQ') but they have not done any research. So, Steve is analyzing other green energy stocks along with DQ to advice his parents. Steve has collected green energy stocks data for years 2017 and 2018 and exploring annual return and total volumes for each stock to make appropriate recommendation to his parents.

### Purpose
Steve has built a VBA script to cature all the stocks in the work sheets with its total volume and annual return. Now, the purpose of this project is to refactor the code Steve has created and improve the performance.

## Results

### Stock Performance between 2017 and 2018:
Using the VBA script we have automated the process to collect the total volume and annual return of each stock for 2017 and 2018. Taking a quick peek at the results we see that most of the stocks have fared well in 2017 but not so good in 2018.

-- pic 2017 and 2018 side by side

First, lets take a look at the stock 'DQ' because Steve's parents were interested in that. DQ has outperformed the other stocks in 2017, gained 199.4% market value. However it is the worst perfrorming stock in 2018, lost 62.6% of the market value. 
There are two stocks which performed well in a relatively unpleasnt market for green energy stocks in 2018, they are *ENPH* and *RUN*. Looking at the previous year's performance for ENPH and RUN, ENPH has done better in 2017.
So, we can conclude that *ENPH* is a better stock to invest in.

### Execution times of Original script and Refactored script:
One of the advantages of refactoring code is the improved performance. To refactor the code, we are using arrays to store the output values, tickerVolumes, tickerStartingPrices and tickerEndingPrices.
The difference between the original code and refactored code is minor but there is slight improvement in performance. Below are the execution times for original script and refactored script

-- pic 2017 and 2018 side by side - original script
-- pic 2017 and 2018 side by side - refactored script

## Summary
  
    * What are the advantages or disadvantages of refactoring code?
      Advantages:
      
      Disadvantages:
      
    * How do these pros and cons apply to refactoring the original VBA script?
    
