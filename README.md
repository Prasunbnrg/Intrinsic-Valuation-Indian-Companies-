 The code can be used to calculate Intrinsic Valuation of **Non-Financial Service Firms** listed company in National Stock Exchange of India (NSE) and Bombay Stock Exchange (BSE). 
The code takes financial data from from *Money Control* and historical data is taken from *Yahoo Finance*. 
A minumun three (3) years financial data needs to be available on Money Control for successfull execution of the code. 

# Input /Output results  
The file path of the excel file *Valuation_InputOutputSheet_R0.xlsx* has to be given as input to the python code. The code takes input and stores the output in the excel file. 

The excel sheet has **6** sheets.
**Instructions** 
	This sheet explains how to use the file and how to put necessary details for doing intrinsic valuation.
**Input Sheet** 
	This sheet needs the inputs for doing the valuation. Following informations are needed as input for the code
	 	* Company Name
	  	* Company url on Money Control
	  	* Yahoo Finance ticker
	 	* Government 10 year T-bond rates
	 	* Moody's rating of the Country
	 	* Equity Risk Premium of mature market (US)
	 	* Tax Rate (%)
 	 	*  Average duration of debt instruments
	 	* Industry average beta and debt-to-equity ratio
 	 	* Capital Expenditure for past three (3) years
	 	* Geographical-wise revenue distribution
	 	* Assumptions of change in growth in EBIT, capital expenditure, depreciation and amortisation, working capital
	 	* The number of years and period for valuation
**Input Financials** 
	This sheet stores all actual financial numbers for performing the intrinsic valuation
**Output Sheet** 
	This sheet stores the output of the code. Following informations are displayed on this page.
		* Company Name
	  	* Risk free rate
	  	* Country default spread
	 	* Company default spread
	 	* Historical beta
	 	* Bottom-Up beta
		* Equity risk premium
		* Cost of debt
		* Cost of capital
		* WACC
		* EBIT growth 
		* Value of the company
		* Market Capitalisation
		* Shares Outstanding
		* Current market price
		* Value per share
		* Observation (Underpriced/Overpriced)
		* Forecasted free cash flow to firm
**CountryDefaultSpread**
	This sheet shows the country-wise defaults spreads.
	Source: Aswath Damodaran Website (*http://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html*)
**CompanyDefaultSpread**
	This sheet shows a synthetic way for calculating company defualt spreads and company rating from interest coverage ratios.
	Source: Aswath Damodaran Website (*http://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ratings.html*)

Play with the numbers, you might get interesting insights from the code
#IntrinsicValuationwithPython
#HappyCoding
 

 	

