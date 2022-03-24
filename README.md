# DQ Stock Analysis

## Overview of Project

For this project, we are creating VBA code (and refactoring it) to compare stocks from 11 different companies in the years 2017 and 2018.

### Purpose

The purpose of this analysis is to evaluate the returns for DAQO New Energy Corp's stocks for the years 2017 and 2018. This analysis compares other companies to find which company's stocks are the best financial decision for investors. The original VBA code was refactored for time efficiency.

## Results

### Stock Performance
In 2017, all stocks except TERP had a return. In 2018, only ENPH and RUN had returns. Per the purpose of evaluating stocks for investors, the ENPH company was the only company that net returns for both years, making it the most profitable stock. The TERP company had losses for both years, making it an obvious company to avoid investing in. Some companies had huge returns in 2017, and had small losses by comparison in 2018. The best example of this is the SEDG company. In 2017, they had a huge return of 184.5%. Their loss in 2018 was -7.8%, but compared to DQ's loss of -62.6%, SEDG investors would have a larger profit margin overall than those invested in the DQ stocks.

### Refactored Code
To increase efficiency, the original VBA code was refactored to run the analysis macros in less time.
A key element to refactoring is creating a structure to guide the new and improved code. Making comments within the code are integral to this structure.
This is the orignial code run time for the 2017 stock analysis: ![VBA_Challenge_2017 Original Code](https://user-images.githubusercontent.com/96644316/159818057-43ddd5bd-0c8e-4111-b77c-3d95956ccf03.png)
This is the refactored code run time for the 2017 stock analysis: ![VBA_Challenge_2017](https://user-images.githubusercontent.com/96644316/159818066-5f48065c-c110-4907-8ecc-17497a230307.png)
This is the orignial code run time for the 2018 stock analysis: ![VBA_Challenge_2018 Original Code](https://user-images.githubusercontent.com/96644316/159818133-5cde5f7d-0aa2-4a4a-8a10-9dbd98d6a409.png)
This is the refactored code run time for the 2017 stock analysis: ![VBA_Challenge_2018](https://user-images.githubusercontent.com/96644316/159818156-b385edac-7154-46ea-8d4f-bc2a1a8e3ad6.png)
The refactored code was much faster for both analyses.

## Summary

### -What are the advantages or disadvantages of refactoring code?
A few advantages of refactoring code is that it increases efficiency by restructuring the code to be more consice, to run faster, and to be easier to read (especially when working on a project with others!). A disadvantage of refactoring code is that it can be very time consuming to refactor depending on how large the data set is, and what analyses need to be ran.

### -How do these pros and cons apply to refactoring the original VBA script?
Long term, a pro to refactoring is that continuously refactored code is code that will stay up-to-date and easier to improve. A con to refactoring code is the chance that you may change a part of the code and "break" something- thus spending more time trying to fix what you've broken.
In our instance here, the refactoring helped improve macro run time (as seen in the screenshots from the resources folder, or under the previous "Refactored Code" heading). For such a small data set, the time set aside to refactor the code likely did not justify the time saved for the macro run time since the run times are only milliseconds apart.
