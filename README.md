# the-VBA-of-Wall-Street
Excel VBA scripting on stock market data.

Contents:
**/code** - contains VBA code to summarize stock ticker data

**/data** - contains Excel data to test code on (*note* contains 100+MB of data)

**/images** - contains screenshots of solution in Excel

Directions: This submission solves the Hard Challenge level. Copy the macro code over to Excel and run the HardChallenge() macro to test the solution.

Additionally, this project uses VBA scripting in Excel to summarize sample stock market data according to three different levels:

1. Easy - Returns the following using Easy(), EasyChallenge() macros: 
    * stock ticker symbol
    * total stock volume per ticker
![Easy Solution](./images/1-easy/easy2014.png)
2. Moderate - Returns the following using Moderate(), ModerateChallenge() macros: 
    * stock ticker symbol
    * yearly change from opening price at the beginning of a given year to the closing price at the end of that year
    * percent change from opening price at the beginning of a given year to the closing price at the end of that year
    * total stock volume per ticker
\
    *NOTE*: if any stock contains opening price of 0, then the percent change is defaulted to NULL value in the corresponding sumamry cell since we cannot divide by 0
![Moderate Solution](./images/2-moderate/moderate2014.png)
3. Hard - Returns the following using Hard(), HardChallenge() macros:
    * Contains everything from the moderate level
    * stock with "Greatest % increase"
    * stock with "Greatest % decrease"
    * stock with "Greatest total volume"
![Hard Solution](./images/3-hard/hard2014.png)

Also note that there are "Clear`[SolutionLevel`]()" macros included to easily clear out cells and formatting for a clean re-run of VBA summary scripts if needed.

The *Challenge() macros loop through each sheet in the Excel workbook and applies the solution to every sheet.