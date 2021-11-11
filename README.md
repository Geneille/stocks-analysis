# Stocks_Analysis

# Overview of Project: 

Steve wants to find the total daily volume and yearly return for each stock from the year 2017 and 2018. While Steve can perform the analysis in excel, he prefers and requires the analysis to be automated, which would therefore not only be applicable for the current stock data but also extendable for more data of the same type. A code is written in VBA can perform such an analysis. VBA is used to automate tasks and perform several other functions including creating and organizing spreadsheets. Further, the code should be refractured to improve its design, structure and runtime, while preserving its functionality.
  

  ## Aim

* Write a code in VBA to calulate total daily volume and yearly return for each stock 
from user defined year input.

* Refracture the code to reduce the runtime so  that the analysis is performed in the 
shortest possible time

# Analysis & Results

  An overview of the analysis is briefly outlined below (see VBA_Challenge.xlms for a more detailed view of the code) 

   1. Variables, output worksheet and table header (for table) to store output data were first defined 

   2. The 'InputBox' function in vba was utilized at the beginning of the code to allow the user to state which year data should be analyzed. This is shown in the image below

![image](https://user-images.githubusercontent.com/92636438/141217539-7c685d74-4667-4555-8a1e-cafa6e3f8eb8.png)

   3. An array (called 'tickers') was created to recall the different tickers

   4. Three output arrays was created to store the required data for each ticker

![image](https://user-images.githubusercontent.com/92636438/141219318-60f9ae88-d705-434d-859b-88a84cf5be5d.png)


   5. A variable called 'tickerIndex', defined globally, was created to actively move through the code for each array. 'tickerIndex' was initalized to 0 as this is the first element in any array.
   
   ![image](https://user-images.githubusercontent.com/92636438/141217089-09abceef-48b9-447d-b2e6-8f9f7a9b70a9.png)


   6. A 'for' loop was used to loop over all the rows in the spreadsheet to access the required information to complete the necessary calculation for the desired output. 
   
   ![image](https://user-images.githubusercontent.com/92636438/141217769-f5c59841-169b-4df7-b43e-31b071278bfe.png)

   
   7. Code to perform the necessary caculations for the different ouput variables is depicted in the images below, as well as the code/condition that ensures the computations is being performed (correctly) for the specific ticker.

![image](https://user-images.githubusercontent.com/92636438/141217321-f7c82f4c-c36f-473e-bafa-cc4964f9c8ad.png)
![image](https://user-images.githubusercontent.com/92636438/141217949-b4f0b20e-0876-4edc-be45-c297f03f459b.png)






   8. Finally, code to store the computed output and format the output information was written

![image](https://user-images.githubusercontent.com/92636438/141218359-abc7a2f6-e765-4e45-bb9b-31b0e3e4030b.png)



   9. The code for execution time for the analysis to be performrd was calculated using the code below
 ![image](https://user-images.githubusercontent.com/92636438/141218568-6c86a6bc-fda3-4a5c-86f1-96d3a7e6040f.png)





