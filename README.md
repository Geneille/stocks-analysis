# Stocks_Analysis

# Overview of Project: 

Daily volume and open and ending prices for a set of stocks for the year 2017 and 2018 was collected. The total daily volume and yearly return for each stock from both years are to be calulated  to gain insight into the profitability of these stocks. While the analysis can be performed in excel, it is a requirement that the analysis is to be automated, which would therefore not only be applicable for the current stock data but also extendable for more data of the same type. A code written in VBA can perform such an analysis. VBA is used to automate tasks and perform several other functions including creating and organizing spreadsheets. Further, the code is to be refractured to improve its design, structure and runtime, while preserving its functionality.
  

   ## Aim

* Write a code in VBA to calulate total daily volume and yearly return for each stock from user defined year input.

* Refracture the code to reduce the runtime so that the analysis is performed in the shortest possible time.

# Analysis & Results

  An overview of the analysis is briefly outlined below. 

   1. Variables, output worksheet and table header (for table) to store output data were first defined.

<img width="400" alt="Screenshot 2021-11-13 051925" src="https://user-images.githubusercontent.com/92636438/141615083-b9011d63-7ddf-4633-aaf3-2ae0d229bcd3.png">


   2. The 'InputBox' function in vba was utilized at the beginning of the code to allow the user to state which year data should be analyzed. This is shown in the image below

![image](https://user-images.githubusercontent.com/92636438/141217539-7c685d74-4667-4555-8a1e-cafa6e3f8eb8.png)

   3. An array (called 'tickers') was created to recall the different tickers.

   4. Three output arrays was created to store the required data for each ticker.

![image](https://user-images.githubusercontent.com/92636438/141219318-60f9ae88-d705-434d-859b-88a84cf5be5d.png)


   5. A variable called 'tickerIndex', defined globally, was created to actively move through the code for each array. 'tickerIndex' was initalized to 0 as this is the first element in any array.
   
   ![image](https://user-images.githubusercontent.com/92636438/141217089-09abceef-48b9-447d-b2e6-8f9f7a9b70a9.png)


   6. A 'for' loop was used to loop over all the rows in the spreadsheet to access the required information to complete the necessary calculation for the desired output. 
   
   ![image](https://user-images.githubusercontent.com/92636438/141217769-f5c59841-169b-4df7-b43e-31b071278bfe.png)

   
   7. Code to perform the necessary caculations for the different ouput variables is depicted in the images below, as well as the code/condition that ensures the computations is being performed (correctly) for the specific ticker.

<img width="416" alt="Screenshot 2021-11-13 052805" src="https://user-images.githubusercontent.com/92636438/141615243-4121bc19-c4fe-4335-9b7f-15285bdb70ca.png">



   8. Finally, code to store the computed output and format the output information was written

<img width="416" alt="Screenshot 2021-11-13 053447" src="https://user-images.githubusercontent.com/92636438/141615421-8234613b-68c6-458d-8f97-1faf59fbfa99.png">


   9. The code for execution time for the analysis to be performrd was calculated using the code below
   
<img width="525" alt="Screenshot 2021-11-13 054317" src="https://user-images.githubusercontent.com/92636438/141615564-a2a02aa0-c65c-44b8-a902-3030c4ca37cb.png">

 

  ## Results
  
  The results for the year 2017 and the runtime is depicted below
  
  <img width="150" alt="Screenshot 2021-11-08 172635" src="https://user-images.githubusercontent.com/92636438/141614859-e94e12b8-a647-4713-9c31-8e8e442e25fc.png">

<img width="306" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/92636438/141614881-cb7aca19-5487-478a-8c7c-4172b13c86ae.png">

 
 The results for the year 2018 and the runtime is depicted below
 
 <img width="150" alt="Screenshot 2021-11-08 172803" src="https://user-images.githubusercontent.com/92636438/141614952-3c15d911-fba8-4af9-908f-33785b2f29ab.png">

<img width="311" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/92636438/141614964-0a523b8d-3add-408b-b083-c7288cd3fe51.png">


# Summary

Code refactoring aims to improve the overall capability of the code without affecting its functionality/purpose. In this regard, the code still produces the same result as it was originally design for, but the code gains the following advantages: its design is improved which helps programming faster, it is easier to understand and read, it is easier to maintain which may help finding bugs faster, and it is more extensible and flexible as it is easier to add many other functions and be more adaptive. On the other hand, refactoring can be both cost and time ineffective. Refactoring can be time consuming as you try to rewrite parts of the code perhaps by introducing global variables or reducing the number of loops, which may introduce errors and mistakes which would therefore require more time finding and fixing.

For this project, the original VBA script worked flawlessly without any errors, but it had a longer runtime for the results output. By introducing a global variable to run through the code, the number of lops in the code was reduced and the refactored code programming was faster as the runtime for the output results was shorter. However, it was very time consuming as a lot of time was spent debugging the code. A lot of time was spent testing different parts of the code to ensure that that specific element was producing the correct results, for example the quotient for the ‘return’ was not being divided by zero. Specifically for this code, a ‘Runtime error’ code 6 - overflow was the main feedback error. When an attempt was made to fix it another error was sometimes introduced which had to be reset. Finally, it was observed the variable (‘tickerIndex’) that was running through the arrays, that had 12 elements, was overshooting beyond the number of elements in the array and was continuously incrementing. The bug was finally fixed and the runtime for the refactored code was shorter than the original.


