# VBA_Challenge
Challenge_Week2_Laura_Roa
This repository contains a list of stocks by ticker by year, with dayli open and close values, as well as volumes. 
Once you run the macros, you will be able to get 2 summary tables on each tab. Each year has a dedicated tab.
   
   
   Instructions to run macros:
In Visual Basic menu, you will find 3 different modules to run. Please, run them in order (Part1, Part2, and Part3)

Part1. First module will produce titles for a summary page, will get the unique ticker names and their total volume.
   Note: There was an issue whe getting the total volume per ticket, as the total volume often exceeded the range allowance of double variable. I found LongPtr function for long numbers on https://learn.microsoft.com/.
   
Part2. Second module will calculate the yearly change (Dec 31 compared to Jan 2) variance as well as the percentage of variance (same period).
   
Part3. Third module will calculate a second summary page that has the greater values of all 3 columns (year change, variance in percentage and total volume).
   Note: The column "K" should have a % format. I found the code for it: Format(percent_change, "#.##%") and I indicated the name of the website on my VB comments.

Due to my lack of expertise in working on one terminal all at once, I decided to take it into parts, so that I could make sure things were working correctly and if they weren't, then I could address the issue easily.
If I were to have more time, I could have figured it out, but I believe you will be satisfied with the results.

Please, note that each page of the worksheet has a given name which is a year (2018, 2019, etc). I used each tab name in a text format, and put the start date as :Ws name + "0102", and end date as: Ws name + "1231". This way will allow the macro to run and work in every single sheet. 

There are 3 files in my repository that has all 3 modules codes, as well as a word document for you to see the final product per tab, as requested on the challenge.

