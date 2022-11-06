## Table of Contents
1. [General Info](#general-info)
2. [Installation](#installation)
3. [Requirements](#requirements)
4. [Collaboration](#collaboration)
5. [References](#references)
### General Info
Project Status:  **Complete**

This is a script used to summarize stock data, looking at year over year changes in stock price, 
the percent difference between the year opening and closing values, and total volume.  It also 
summarizes the greatest and worst performing stock in the year, and the stock with the greatest 
volume of trades.

It will perform the analysis on any number of worksheets in the workbook

### Screenshot
Example screenshots included in repository

## Installation
Download and import script into Excel sheet

The code requires the following data in columns in the following order:
ticker, date, open, high, low, close, vol

These columns must be sorted ascending first by ticker, then by date.

## Requirements
    Create a script that loops through all the stocks for one year and outputs the following information:

    * The ticker symbol.
    * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    * The total stock volume of the stock.

    **Note:** Make sure to use conditional formatting that will highlight positive change in green and negative change in red.
    The result should match the following image:

    ![moderate_solution](Images/moderate_solution.png)

    **Bonus**   
    Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total 
    volume". Make the appropriate adjustments to your VBA script to allow it to run on every worksheet (that is, every year) just 
    by running the VBA script once.  The solution should match the following image:

    ![hard_solution](Images/hard_solution.png)

    **Other Considerations**
    * Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller and will allow you to test 
    faster. Your code should run on this file in less than 3 to 5 minutes.  ** Make sure that the script acts the same on every 
    sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with one click of a button.

    * Some assignments, like this one, contain a bonus. It is possible to achieve proficiency for this assignment without completing 
    the bonus. The bonus is an opportunity to further develop your skills and be rewarded extra points for doing so.

    **Submission**
    To submit, please upload the following to GitHub:
    * A screen shot for each year of your results on the multi-year stock data.
    * VBA scripts as separate files.
    Be sure to commit regularly to your repository and that it contains a README.md file.  After saving your work, create a 
    shareable link and submit the link to <https://bootcampspot-v2.com/>.

    **Rubric**
    [Unit 2 Rubric - VBA Homework - The VBA of Wall Street](https://docs.google.com/document/d/1OjDM3nyioVQ6nJkqeYlUK7SxQ3WZQvvV3T9MHCbnoWk/edit?usp=sharing)

## Collaboration
Feel free to use, and send feedback to me regarding any changes/upgrades/features that should be added.

## References
U of T Bootcamp instructors, instruction and exercises
www.wallstreetmojo.com
www.automateexcel.com
www.Exceldemy.com
www.stackoverflow.com
www.microsoft.com
www.ionos.com

|:--------------|:-------------:|--------------:|
| text-align left | text-align center | text-align right |