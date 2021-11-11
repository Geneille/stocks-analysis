# Stocks_Analysis

# Overview of Project: 

Steve wants to find the total daily volume and yearly return for each stock from the year 2017 and 2018. While Steve can perform the analysis in excel, he prefers and requires
the analysis to be automated, which would therefore not only be applicable for the current stock data but also extendable for more data of the same type. A code is written in VBA can perform such an analysis. VBA is used to automate tasks and perform several  other functions including creating and organizing spreadsheets. Further, the code should 
be refractured to improve its design, structure and runtime, while preserving its functionality.
  

## Aim

*Write a code in VBA to calulate total daily volume and yearly return for each stock 
from user defined year input.

* Refracture the code to reduce the runtime so  that the analysis is performed in the 
shortest possible time

# Analysis & Results

	An overview of the analysis is briefly outlined below

	1. Variables, output worksheet and table header (for table) to store output data were first 
defined 
	2. The 'InputBox' function in vba was utilized at the beginning of the code 
to allow the user to state which year data should be analyzed. This is shown in the image below
	3. An array (called 'tickers') was created to recall the different tickers
	4. Three output arrays was created to store the required data for each ticker
	5. A variable called 'tickerIndex', defined globally, was created to actively move
through the code for each array. 'tickerIndex' was initalized to 0 as this is the first element in any array
	6. A 'for' loop was used to loop over all the rows in the spreadsheet to access
the required information to complete the necessary calculation for the desired output.
Code to perform the necessary caculations for the different ouput variables is depicted in
the images below, as well as the code/condition that ensures the co computions is 
being performed (correctlyz) for the specific ticker.
	7. The code to increment the tickerIndex, shown below, was then written
	8. Code to store the computed output was written
	9. Finally code to format the output information was written 

	
 




